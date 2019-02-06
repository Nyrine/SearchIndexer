Imports System.IO
Public Class addin

    Public Function _getDriveType(ByVal DriveType As Integer) As String
        Select Case DriveType
            Case 0
                Return "Unknown"
            Case 1
                Return "NoRootDirectory"
            Case 2
                Return "Removable"
            Case 3
                Return "Fixed"
            Case 4
                Return "Network"
            Case 5
                Return "Optical"
            Case 6
                Return "RAM"
            Case Else
                Return "Error not Known"
        End Select
    End Function
End Class
