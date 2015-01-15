Imports System
Imports System.ComponentModel.Design

Class Globals

    Public Const PackageGuidString As String = "d3a46f6c-a314-4c66-af21-5e6a0d4ba1a0"

    Public Const CommandSetGuidString As String = "210b9da1-9778-4aac-88db-523661b6c4b8"
    Public Shared ReadOnly CommandSetGuid As New Guid(CommandSetGuidString)

    ' Command to bring up KeyBindings dialog and to cycle through categories (if i add this)
    Public Const CommandID As UInteger = &H100
    Public Shared ReadOnly KeyBindingsCommand As New CommandID(Globals.CommandSetGuid, CInt(Globals.CommandID))

    ' Command to cycle backwards through categories (if i add this)
    Public Const CommandBackID As UInteger = &H101
    Public Shared ReadOnly KeyBindingsBackCommand As CommandID = New CommandID(Globals.CommandSetGuid, Globals.CommandBackID)

    'Public Const guidToolWindowPersistanceString As String = "f7539fa7-dda1-4c70-a7dc-11f125a63092"
    'Public Const FeedbackLink As String = "http://go.microsoft.com/fwlink/?LinkID=TODO"

End Class