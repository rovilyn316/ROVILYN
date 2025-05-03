Imports System.Data.Odbc
Public Class Class2
    Private db As New vbHelperClass()

    Public Function getCourses() As DataTable
        Dim query As String = "select * from tblCourses"
        Return db.ExecuteQuery(query)
    End Function
End Class
