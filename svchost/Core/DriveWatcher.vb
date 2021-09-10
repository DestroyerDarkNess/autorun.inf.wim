' ***********************************************************************
' Author   : Elektro
' Modified : 11-November-2015
' ***********************************************************************
' <copyright file="DriveWatcher.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace Core.Engine.Watcher

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' A device insertion and removal monitor.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Class DriveWatcher : Inherits NativeWindow : Implements IDisposable

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the connected drives on this computer.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Drives As IEnumerable(Of DriveInfo)
            <DebuggerStepThrough>
            Get
                Return DriveInfo.GetDrives
            End Get
        End Property

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value that determines whether the monitor is running.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property IsRunning As Boolean
            <DebuggerStepThrough>
            Get
                Return Me.isRunningB
            End Get
        End Property
        Private isRunningB As Boolean

#End Region

#Region " Events "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' A list of event delegates.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private ReadOnly events As EventHandlerList

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Occurs when a drive is inserted, removed, or changed.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Custom Event DriveStatusChanged As EventHandler(Of DriveStatusChangedEventArgs)

            <DebuggerNonUserCode>
            <DebuggerStepThrough>
            AddHandler(ByVal value As EventHandler(Of DriveStatusChangedEventArgs))
                Me.events.AddHandler("DriveStatusChangedEvent", value)
            End AddHandler

            <DebuggerNonUserCode>
            <DebuggerStepThrough>
            RemoveHandler(ByVal value As EventHandler(Of DriveStatusChangedEventArgs))
                Me.events.RemoveHandler("DriveStatusChangedEvent", value)
            End RemoveHandler

            <DebuggerNonUserCode>
            <DebuggerStepThrough>
            RaiseEvent(ByVal sender As Object, ByVal e As DriveStatusChangedEventArgs)
                Dim handler As EventHandler(Of DriveStatusChangedEventArgs) =
                    DirectCast(Me.events("DriveStatusChangedEvent"), EventHandler(Of DriveStatusChangedEventArgs))

                If (handler IsNot Nothing) Then
                    handler.Invoke(sender, e)
                End If
            End RaiseEvent

        End Event

#End Region

#Region " Events Data "

#Region " DriveStatusChangedEventArgs "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Contains the event-data of a <see cref="DriveStatusChanged"/> event.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public NotInheritable Class DriveStatusChangedEventArgs : Inherits EventArgs

#Region " Properties "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the device event that occurred.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The drive info.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property DeviceEvent As DeviceEvents
                <DebuggerStepThrough>
                Get
                    Return Me.deviceEventsB
                End Get
            End Property
            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' ( Backing field )
            ''' The device event that occurred.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Private ReadOnly deviceEventsB As DeviceEvents

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Gets the drive info.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <value>
            ''' The drive info.
            ''' </value>
            ''' ----------------------------------------------------------------------------------------------------
            Public ReadOnly Property DriveInfo As DriveInfo
                <DebuggerStepThrough>
                Get
                    Return Me.driveInfoB
                End Get
            End Property
            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' ( Backing field )
            ''' The drive info.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Private ReadOnly driveInfoB As DriveInfo

#End Region

#Region " Constructors "

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Prevents a default instance of the <see cref="DriveStatusChangedEventArgs"/> class from being created.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerNonUserCode>
            Private Sub New()
            End Sub

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Initializes a new instance of the <see cref="DriveStatusChangedEventArgs"/> class.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            ''' <param name="driveInfo">
            ''' The drive info.
            ''' </param>
            ''' ----------------------------------------------------------------------------------------------------
            <DebuggerStepThrough>
            Public Sub New(ByVal deviceEvent As DeviceEvents, ByVal driveInfo As DriveInfo)

                Me.deviceEventsB = deviceEvent
                Me.driveInfoB = driveInfo

            End Sub

#End Region

        End Class

#End Region

#End Region

#Region " Event Invocators "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Raises <see cref="DriveStatusChanged"/> event.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="e">
        ''' The <see cref="DriveStatusChangedEventArgs"/> instance containing the event data.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Protected Overridable Sub OnDriveStatusChanged(ByVal e As DriveStatusChangedEventArgs)

            RaiseEvent DriveStatusChanged(Me, e)

        End Sub

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies a change to the hardware configuration of a device.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/aa363480%28v=vs.85%29.aspx"/>
        ''' <para></para>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/aa363232%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Public Enum DeviceEvents As Integer

            ' *****************************************************************************
            '                            WARNING!, NEED TO KNOW...
            '
            '  THIS ENUMERATION IS PARTIALLY DEFINED TO MEET THE PURPOSES OF THIS PROJECT
            ' *****************************************************************************

            ''' <summary>
            ''' The current configuration has changed, due to a dock or undock.
            ''' </summary>
            Change = &H219

            ''' <summary>
            ''' A device or piece of media has been inserted and becomes available.
            ''' </summary>
            Arrival = &H8000

            ''' <summary>
            ''' Request permission to remove a device or piece of media.
            ''' <para></para>
            ''' This message is the last chance for applications and drivers to prepare for this removal.
            ''' However, any application can deny this request and cancel the operation.
            ''' </summary>
            QueryRemove = &H8001

            ''' <summary>
            ''' A request to remove a device or piece of media has been canceled.
            ''' </summary>
            QueryRemoveFailed = &H8002

            ''' <summary>
            ''' A device or piece of media is being removed and is no longer available for use.
            ''' </summary>
            RemovePending = &H8003

            ''' <summary>
            ''' A device or piece of media has been removed.
            ''' </summary>
            RemoveComplete = &H8004

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Specifies a computer device type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="https://msdn.microsoft.com/en-us/library/windows/desktop/aa363246%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Private Enum DeviceType As Integer

            ' *****************************************************************************
            '                            WARNING!, NEED TO KNOW...
            '
            '  THIS ENUMERATION IS PARTIALLY DEFINED TO MEET THE PURPOSES OF THIS PROJECT
            ' *****************************************************************************

            ''' <summary>
            ''' Logical volume.
            ''' </summary>
            Logical = &H2

        End Enum

#End Region

#Region " Types "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Contains information about a logical volume.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/aa363249%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        <StructLayout(LayoutKind.Sequential)>
        Private Structure DevBroadcastVolume

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' The size of this structure, in bytes.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Size As UInteger

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Set to DBT_DEVTYP_VOLUME (2).
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Type As UInteger

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' Reserved parameter; do not use this.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Reserved As UInteger

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' The logical unit mask identifying one or more logical units.
            ''' Each bit in the mask corresponds to one logical drive.
            ''' Bit 0 represents drive A, bit 1 represents drive B, and so on.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Mask As UInteger

            ''' ----------------------------------------------------------------------------------------------------
            ''' <summary>
            ''' This parameter can be one of the following values:
            ''' '0x0001': Change affects media in drive. If not set, change affects physical device or drive.
            ''' '0x0002': Indicated logical volume is a network volume.
            ''' </summary>
            ''' ----------------------------------------------------------------------------------------------------
            Public Flags As UShort

        End Structure

#End Region

#Region " Constructor "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of <see cref="DriveWatcher"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Sub New()

            Me.events = New EventHandlerList

        End Sub

#End Region

#Region " Public Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Starts monitoring.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Exception">
        ''' Monitor is already running.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Overridable Sub Start()

            If (Me.Handle = IntPtr.Zero) Then
                MyBase.CreateHandle(New CreateParams)
                Me.isRunningB = True

            Else
                Throw New Exception(message:="Monitor is already running.")

            End If

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Stops monitoring.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <exception cref="Exception">
        ''' Monitor is already stopped.
        ''' </exception>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Overridable Sub [Stop]()

            If (Me.Handle <> IntPtr.Zero) Then
                MyBase.DestroyHandle()
                Me.isRunningB = False

            Else
                Throw New Exception(message:="Monitor is already stopped.")

            End If

        End Sub

#End Region

#Region " Private Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the drive letter stored in a <see cref="DevBroadcastVolume"/> structure.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="Device">
        ''' The <see cref="DevBroadcastVolume"/> structure containing the device mask.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The drive letter.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Private Function GetDriveLetter(ByVal device As DevBroadcastVolume) As Char

            Dim driveLetters As Char() = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray

            Dim deviceID As New BitArray(BitConverter.GetBytes(device.Mask))

            For i As Integer = 0 To deviceID.Length

                If deviceID(i) Then
                    Return driveLetters(i)
                End If

            Next i

            Return Nothing

        End Function

#End Region

#Region " Window Procedure (WndProc) "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Invokes the default window procedure associated with this window to process messages for this Window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="m">
        ''' A <see cref="T:System.Windows.Forms.Message"/> that is associated with the current Windows message.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Protected Overrides Sub WndProc(ByRef m As Message)

            Select Case m.Msg

                Case DeviceEvents.Change ' The hardware has changed.

                    If (m.LParam = IntPtr.Zero) Then
                        Exit Select
                    End If

                    ' If it's an storage device then...
                    If Marshal.ReadInt32(m.LParam, 4) = DeviceType.Logical Then

                        ' Transform the LParam pointer into the data structure.
                        Dim currentWDrive As DevBroadcastVolume =
                            DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DevBroadcastVolume)), DevBroadcastVolume)

                        Dim driveLetter As Char = Me.GetDriveLetter(currentWDrive)
                        Dim deviceEvent As DeviceEvents = DirectCast(m.WParam.ToInt32, DeviceEvents)
                        Dim driveInfo As New DriveInfo(driveLetter)

                        Me.OnDriveStatusChanged(New DriveStatusChangedEventArgs(deviceEvent, driveInfo))

                    End If

            End Select

            ' Return Message to base message handler.
            MyBase.WndProc(m)

        End Sub

#End Region

#Region " Hidden methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Serves as a hash function for a particular type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function GetHashCode() As Integer
            Return MyBase.GetHashCode
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the <see cref="System.Type"/> of the current instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' The exact runtime type of the current instance.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function [GetType]() As Type
            Return MyBase.GetType
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Determines whether the specified <see cref="System.Object"/> instances are considered equal.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function Equals(ByVal obj As Object) As Boolean
            Return MyBase.Equals(obj)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Returns a String that represents the current object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function ToString() As String
            Return MyBase.ToString
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Assigns a handle to this window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Sub AssignHandle(ByVal handle As IntPtr)
            MyBase.AssignHandle(handle)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a window and its handle with the specified creation parameters.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Sub CreateHandle(ByVal cp As CreateParams)
            MyBase.CreateHandle(cp)
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Destroys the window and its handle.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Sub DestroyHandle()
            MyBase.DestroyHandle()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Releases the handle associated with this window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Sub ReleaseHandle()
            MyBase.ReleaseHandle()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieves the current lifetime service object that controls the lifetime policy for this instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function GetLifeTimeService() As Object
            Return MyBase.GetLifetimeService
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Obtains a lifetime service object to control the lifetime policy for this instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function InitializeLifeTimeService() As Object
            Return MyBase.InitializeLifetimeService
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Creates an object that contains all the relevant information to generate a proxy used to communicate with a remote object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Function CreateObjRef(ByVal requestedType As Type) As System.Runtime.Remoting.ObjRef
            Return MyBase.CreateObjRef(requestedType)
        End Function

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Invokes the default window procedure associated with this window.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        <DebuggerNonUserCode>
        Public Shadows Sub DefWndProc(ByRef m As Message)
            MyBase.DefWndProc(m)
        End Sub

#End Region

#Region " IDisposable Implementation "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' To detect redundant calls when disposing.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private isDisposed As Boolean

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Releases all the resources used by this instance.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Public Sub Dispose() Implements IDisposable.Dispose

            Me.Dispose(isDisposing:=True)
            GC.SuppressFinalize(obj:=Me)

        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        ''' Releases unmanaged and - optionally - managed resources.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="isDisposing">
        ''' <see langword="True"/>  to release both managed and unmanaged resources; 
        ''' <see langword="False"/> to release only unmanaged resources.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        <DebuggerStepThrough>
        Protected Overridable Sub Dispose(ByVal isDisposing As Boolean)

            If (Not Me.isDisposed) AndAlso (isDisposing) Then

                Me.events.Dispose()
                Me.Stop()

            End If

            Me.isDisposed = True

        End Sub

#End Region

    End Class
End Namespace
