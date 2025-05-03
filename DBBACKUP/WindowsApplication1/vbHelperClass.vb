Imports System.Data.Odbc
Public Class vbHelperClass
    Private con As OdbcConnection ' GUMAGAWA NG CONNCTION KAY DATA BASE 
    Private cmd As OdbcCommand ' palautos kay data base  or sql code
    Private reader As OdbcDataReader ' bumabasa ng resulta ng sql


    Public Sub New()
        con = New OdbcConnection("DSN=dbStudentInformation")
    End Sub

    Public Function openConnection() As Boolean
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
                MsgBox("Connection Successfull", vbInformation, "Database")
                Return True
            End If
        Catch ex As Exception
            MsgBox("Error:" & ex.Message, vbCritical, "Database error")
            Return False
        Finally
            GC.Collect()

        End Try
        Return False
    End Function

    Public Sub closedConnection()
        Try
            If con.State = ConnectionState.Open Then
                con.Close()

            End If
        Catch ex As Exception
            MsgBox("Error:" & ex.Message, vbCritical, "Database error")
        End Try
    End Sub


    Public Function GetConnection() As OdbcConnection
        Return con
    End Function

    Public Function ExecuteQuery(ByVal query As String, Optional ByVal parameters As List(Of OdbcParameter) = Nothing) As DataTable
        Dim dt As New DataTable()
        Try
            If openConnection() Then
                Using cmd As New OdbcCommand(query, con)
                    If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
                        cmd.Parameters.AddRange(parameters.ToArray)

                    End If
                    Using reader As OdbcDataReader = cmd.ExecuteReader
                        dt.Load(reader)
                    End Using
                End Using
                closedConnection()
            End If
        Catch ex As Exception
            MsgBox("error: " & ex.Message, vbVerticalTab, "databases")
        End Try

        Return dt

    End Function
End Class