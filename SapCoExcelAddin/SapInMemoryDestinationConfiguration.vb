' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapInMemoryDestinationConfiguration
    Implements IDestinationConfiguration

    Private availableDestinations As Dictionary(Of String, RfcConfigParameters)

    Public Sub New()
        availableDestinations = New Dictionary(Of String, RfcConfigParameters)()
    End Sub

    Public Function GetParameters(ByVal destinationName As String) As RfcConfigParameters Implements IDestinationConfiguration.GetParameters
        Dim foundDestination As RfcConfigParameters = Nothing
        availableDestinations.TryGetValue(destinationName, foundDestination)
        Return foundDestination
    End Function

    Public Function ChangeEventsSupported() As Boolean Implements IDestinationConfiguration.ChangeEventsSupported
        Return True
    End Function

    Public Event ConfigurationChanged As RfcDestinationManager.ConfigurationChangeHandler Implements IDestinationConfiguration.ConfigurationChanged

    Public Sub AddOrEditDestination(ByVal parameters As RfcConfigParameters)

        Dim name As String = parameters(RfcConfigParameters.Name)
        If availableDestinations.ContainsKey(name) Then
            '' Fire a change event
            '' If Not ConfigurationChanged Is Nothing Then
            ''Always check for null on event handlers... If AddOrEditDestination() gets called before this
            ''instance of InMemoryDestinationConfiguration is registered with the RfcDestinationManager, we
            ''would get a NullReferenceException when trying to raise the event... Stupid concept.
            ''                Why(doesn) 't the .NET framework do this for me?

            Dim EventArgs As New RfcConfigurationEventArgs(RfcConfigParameters.EventType.CHANGED, parameters)
            RaiseEvent ConfigurationChanged(name, EventArgs)
            '' End If
        End If
        '' Replace the current parameters of an existing destination or add a new one
        availableDestinations(name) = parameters
        Dim tmp As String = "Application server"
        Dim isLoadBalancing As Boolean = parameters.TryGetValue(RfcConfigParameters.LogonGroup, tmp)
        If isLoadBalancing Then
            tmp = "Load balancing"
        End If
    End Sub

    '' Removes the destination specified by its name

    Public Sub RemoveDestination(ByVal name As String)
        If availableDestinations.Remove(name) Then
            Console.WriteLine("Successfully removed destination {0}", name)
            ''If Not ConfigurationChanged Is Nothing Then  '' Always check for null
            Console.WriteLine("Firing deletion event for destination {0}", name)
            RaiseEvent ConfigurationChanged(name, New RfcConfigurationEventArgs(RfcConfigParameters.EventType.DELETED))
            ''  End If
        Else
            Console.WriteLine("The destination could not be removed since it does not exist")
        End If
    End Sub

    Public Function getAvailableDestinations() As Dictionary(Of String, RfcConfigParameters)
        getAvailableDestinations = availableDestinations
    End Function

End Class
