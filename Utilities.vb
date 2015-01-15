' Copyright (c) Microsoft Corporation.  All rights reserved.
Imports System.ComponentModel
Imports EnvDTE80
Imports EnvDTE
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop

Friend Class Utilities : Inherits Package

    ''' <summary>
    ''' Get the Visual Studio DTE object 
    ''' This allows controls to interact with Visual Studio through the DTE object model.
    ''' </summary>
    Public Shared Function GetDTE() As DTE2
        Return TryCast(Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(GetType(EnvDTE.DTE)), DTE2)
    End Function

    ''' <summary>
    ''' Return the Visual Studio IServiceProvider interface.
    ''' This allows controls to interact with Visual Studio through proffered services.
    ''' </summary>
    ''' <param name="dte">Visual Studio DTE object</param>
    Public Shared Function GetServiceProvider(ByVal dte As DTE2) As ServiceProvider
        Return New ServiceProvider(CType(dte, Microsoft.VisualStudio.OLE.Interop.IServiceProvider))
    End Function
End Class