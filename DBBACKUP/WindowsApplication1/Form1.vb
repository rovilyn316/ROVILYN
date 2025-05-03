Public Class Form1
    Private userDal As New Class3

    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Dim isLoaded As Boolean = userDal.loadUser(dgvList)
        If Not isLoaded Then
            MsgBox("No record Founded", vbInformation, "Databeses")

        End If
    End Sub