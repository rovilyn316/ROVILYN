Imports System.Data.Odbc
Public Class helperClass
    Private con As OdbcConnection
    Private cmd As OdbcCommand
    Private reader As OdbcDataReader

    Public Sub New()
        con = New OdbcConnection("DSN=dbStudentInformation")

    End Sub

    Public Function OpenConnection() As Boolean
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
                MsgBox("Connection Successfull", vbInformation, "Databases")
                Return True
            End If
        Catch ex As Exception
            MsgBox("error:" & ex.Message, vbCritical, "Databases error")
            Return False
        Finally
            GC.Collect()

        End Try
        Return False
    End Function
    Public Sub Connection()
        Try
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MsgBox("Error:" & ex.Message, vbCritical, "Databases error")
        End Try
    End Sub


    Public Sub CloseConnection()
        Try
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MsgBox("Error:" & ex.Message, vbCritical, "Database Error")

        End Try
    End Sub

    Public Function GetConnection() As OdbcConnection
        Return con
    End Function

    Public Function ExecuteQuery(ByVal query As String, Optional ByVal parameters As List(Of OdbcParameter) = Nothing) As DataTable
        Dim dt As New DataTable()
        Try
            If OpenConnection() Then

                Using cmd As New OdbcCommand(query, con)
                    If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
                        cmd.Parameters.AddRange(parameters.ToArray)

                    End If
                    Using reader As OdbcDataReader = cmd.ExecuteReader

                    End Using
                End Using
                CloseConnection()

            End If
        Catch ex As Exception
            MsgBox("Error:" & ex.Message, vbCritical, "Database")
        End Try
        Return dt
    End Function
End Class
