Imports System.IO

Module Infos

    Public Function GetDriveLabel(ByVal dName As String) As String
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo
        For Each d In allDrives
            If d.Name = DName Then
                Return d.VolumeLabel
            End If
        Next
        Return "Error"
    End Function
End Module

