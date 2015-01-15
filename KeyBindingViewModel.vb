Imports System.Collections.ObjectModel
Imports EnvDTE80
Imports System.ComponentModel
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio

Public Class KeyBindingViewModel
    Implements INotifyPropertyChanged, IVsSelectionEvents

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
    Public Event InvokeCommand(ByVal command As String)

    Public DTE As DTE2
    Public IsDataLoaded As Boolean = False

    Private _keybindings As List(Of CommandTableEntry)
    Private _scopedKeyBindings As List(Of CommandTableEntry)
    Private _keyBindingTableNames As Dictionary(Of Guid, String)
    Private _activeKeyBindingTableName As String = "Global"
    Private _inheritedKeyBindingTableName As String = "Global"

    Public Property CurrentDocWindowFrame As IVsWindowFrame

    Public Sub New()

        _keybindings = New List(Of CommandTableEntry)
        _scopedKeyBindings = New List(Of CommandTableEntry)
        _keyBindingTableNames = LoadKeybindingTableNames()

        IsDataLoaded = False

        Dim monsel As IVsMonitorSelection = CType(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(GetType(SVsShellMonitorSelection)), IVsMonitorSelection)
        Dim pdwCookie As UInteger
        Dim returnValue As Integer

        returnValue = monsel.AdviseSelectionEvents(Me, pdwCookie)


    End Sub


    Public Property KeyBindings As List(Of CommandTableEntry)
        Get
            Try

                _scopedKeyBindings = _keybindings.Where(Function(cte As CommandTableEntry)
                                                            If cte.Scope = "Global" Or
                                                                cte.Scope = _activeKeyBindingTableName Or
                                                                cte.Scope = _inheritedKeyBindingTableName Then
                                                                Return True
                                                            Else
                                                                Return False
                                                            End If
                                                        End Function).OrderBy(Function(cte As CommandTableEntry)
                                                                                  Return cte.Scope
                                                                              End Function).ThenBy(Function(cte As CommandTableEntry)
                                                                                                       Return cte.CommandName
                                                                                                   End Function).ToList

                Return _scopedKeyBindings

            Catch ex As Exception
                Return Nothing
            End Try

        End Get
        Set(value As List(Of CommandTableEntry))
            _keybindings = value
            NotifyPropertyChanged("KeyBindings")
        End Set
    End Property

    Public ReadOnly Property FilteredKeyBindings(ByVal filterString As String) As List(Of CommandTableEntry)

        Get
            Try
                If filterString <> "" Then


                    Return _scopedKeyBindings.Where(Function(cte As CommandTableEntry)
                                                        If cte.CommandName.ToUpper.Contains(filterString.ToUpper) Or
                                                            cte.Scope.ToUpper.Contains(filterString.ToUpper) Or
                                                            cte.KeyBinding.ToUpper.Contains(filterString.ToUpper) Then
                                                            Return True
                                                        Else
                                                            Return False
                                                        End If
                                                    End Function).ToList
                Else
                    Return _scopedKeyBindings
                End If
            Catch ex As Exception
                Return Me.KeyBindings
            End Try

        End Get

    End Property

    Public Property ActiveKeyBindingTableName As String
        Get
            Return _activeKeyBindingTableName
        End Get
        Set(value As String)
            If value IsNot Nothing Then
                If value <> _activeKeyBindingTableName Then
                    _activeKeyBindingTableName = value
                    NotifyPropertyChanged("ActiveKeyBindingTableName")
                    NotifyPropertyChanged("KeyBindings")
                End If
            End If
        End Set
    End Property

    Public Property InheritedKeyBindingTableName As String

        Get
            Return _inheritedKeyBindingTableName
        End Get
        Set(value As String)
            If value IsNot Nothing Then
                If value <> _inheritedKeyBindingTableName Then
                    _inheritedKeyBindingTableName = value
                    NotifyPropertyChanged("InheritedKeyBindingTableName")
                    NotifyPropertyChanged("KeyBindings")
                End If
            End If
        End Set
    End Property

    Public Sub LoadModel()

        Dim keys As System.Array
        Dim name As String
        Dim s() As String
       
        For i As Integer = 1 To DTE.Commands.Count

            keys = CType(DTE.Commands.Item(i).Bindings, System.Array)

            If keys.Length > 0 Then
                For j As Integer = 0 To keys.Length - 1
                    name = DTE.Commands.Item(i).Name
                    If String.IsNullOrEmpty(name) = False Then
                        s = keys(j).ToString.Split(":"c)
                        _keybindings.Add(New CommandTableEntry With {.CommandName = name, .Scope = s(0), .KeyBinding = s(2)})
                    End If
                Next j
            End If

        Next i

        Call LoadActiveKeyBindingScopes()
        IsDataLoaded = True

    End Sub

    Friend Sub LoadActiveKeyBindingScopes()
        Dim cmdUIGuid As Guid = Guid.Empty
        Dim inheritKeyBindings As Guid = Guid.Empty

        GetActiveKeyBindingScopes(cmdUIGuid, inheritKeyBindings)

        If Not _keyBindingTableNames.TryGetValue(cmdUIGuid, ActiveKeyBindingTableName) Then
            ActiveKeyBindingTableName = "Global"
        End If

        If Not _keyBindingTableNames.TryGetValue(inheritKeyBindings, InheritedKeyBindingTableName) Then
            InheritedKeyBindingTableName = "Global"
        End If

    End Sub

    Private Sub GetActiveKeyBindingScopes(ByRef cmdUIGuid As Guid, ByRef inheritKeyBindings As Guid)
        Dim monsel As IVsMonitorSelection = CType(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(GetType(SVsShellMonitorSelection)), IVsMonitorSelection)
        Dim objCurDocWindow As Object = Nothing

        Dim hr As Integer = monsel.GetCurrentElementValue(CType(Microsoft.VisualStudio.VSConstants.VSSELELEMID.SEID_DocumentFrame, UInteger), objCurDocWindow)

        If ErrorHandler.Succeeded(hr) And objCurDocWindow IsNot Nothing Then
            CurrentDocWindowFrame = CType(objCurDocWindow, IVsWindowFrame)
            If CurrentDocWindowFrame IsNot Nothing Then
                CurrentDocWindowFrame.GetGuidProperty(CType(__VSFPROPID.VSFPROPID_CmdUIGuid, Integer), cmdUIGuid)
                CurrentDocWindowFrame.GetGuidProperty(CType(__VSFPROPID.VSFPROPID_InheritKeyBindings, Integer), inheritKeyBindings)
            End If
        End If
    End Sub

    Private Function LoadKeybindingTableNames() As Dictionary(Of Guid, String)

        Dim keyBindingTableNames As Dictionary(Of Guid, String)
        Dim hr As Integer = VSConstants.S_OK

        Dim settingsManager As IVsSettingsManager = TryCast(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(GetType(SVsSettingsManager)), IVsSettingsManager)
        Dim settings As IVsSettingsStore = Nothing
        ErrorHandler.ThrowOnFailure(settingsManager.GetReadOnlySettingsStore(CType(__VsSettingsScope.SettingsScope_Configuration, UInteger), settings))

        Dim rootKey As String = "KeyBindingTables"
        Dim shell As IVsShell = TryCast(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(GetType(SVsShell)), IVsShell)
        Dim count As UInteger = 0

        ErrorHandler.ThrowOnFailure(settings.GetSubCollectionCount(rootKey, count))

        keyBindingTableNames = New Dictionary(Of Guid, String)


        For i As UInteger = 0 To count - 1
            Dim keyName As String = ""

            ErrorHandler.ThrowOnFailure(settings.GetSubCollectionName(rootKey, i, keyName))
            Dim keyBindingTableGuid As Guid = New Guid(keyName)

            Dim displayName As String = ""
            Dim packageID As String = ""
            Dim subKeyName As String = rootKey & "\\" & keyName

            settings.GetString(subKeyName, "", displayName)
            settings.GetString(subKeyName, "Package", packageID)
            Dim guidPackage As Guid = New Guid(packageID)

            If displayName(0) = "#" Then
                Dim resID As Integer = 0
                If Integer.TryParse(displayName.Substring(1), resID) Then
                    hr = shell.LoadPackageString(guidPackage, CType(resID, UInteger), displayName)
                End If
            End If

            keyBindingTableNames.Add(keyBindingTableGuid, displayName)
        Next

        Return keyBindingTableNames

    End Function

    Public Overridable Sub NotifyPropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Function OnCmdUIContextChanged5(dwCmdUICookie As UInteger, fActive As Integer) As Integer Implements VisualStudio.Shell.Interop.IVsSelectionEvents.OnCmdUIContextChanged
        LoadActiveKeyBindingScopes()
        Return VSConstants.S_OK
    End Function

    Public Function OnElementValueChanged(elementid As UInteger, varValueOld As Object, varValueNew As Object) As Integer Implements VisualStudio.Shell.Interop.IVsSelectionEvents.OnElementValueChanged
        Return VSConstants.S_OK
    End Function

    Public Function OnSelectionChanged(pHierOld As VisualStudio.Shell.Interop.IVsHierarchy,
                                       itemidOld As UInteger,
                                       pMISOld As VisualStudio.Shell.Interop.IVsMultiItemSelect,
                                       pSCOld As VisualStudio.Shell.Interop.ISelectionContainer,
                                       pHierNew As VisualStudio.Shell.Interop.IVsHierarchy,
                                       itemidNew As UInteger,
                                       pMISNew As VisualStudio.Shell.Interop.IVsMultiItemSelect,
                                       pSCNew As VisualStudio.Shell.Interop.ISelectionContainer) As Integer Implements VisualStudio.Shell.Interop.IVsSelectionEvents.OnSelectionChanged
        Return VSConstants.S_OK
    End Function

End Class