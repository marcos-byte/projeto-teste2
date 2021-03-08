Public Class Form1
	Private Sub BtnTeste_Click(sender As Object, e As EventArgs) Handles BtnTeste1.Click

		Try

			MsgBox("Olá")

		Catch ex As Exception

		End Try

	End Sub

	Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		Try

			TxtTeste1.Text = "Banana"

		Catch ex As Exception

		End Try

	End Sub

	Private Sub BtnTeste2_Click(sender As Object, e As EventArgs) Handles BtnTeste2.Click

		Try

			MsgBox("Tudo bem?")

		Catch ex As Exception

		End Try

	End Sub

	Private Sub MenuItem21_Click(sender As Object, e As EventArgs) Handles MenuItem21.Click

		Try

			MsgBox("A verdade!")

		Catch ex As Exception

		End Try

	End Sub

	Private Sub cmbteste1_Click(sender As Object, e As EventArgs) Handles cmbteste1.Click

		Dim Ponto As Point
		Dim RPA As New ClRpa

		Try

			Ponto = RPA.GetMouse()

			MsgBox(Ponto.X & "|" & Ponto.Y)


		Catch ex As Exception

		End Try

	End Sub

	Private Sub TxtTeste1_Click(sender As Object, e As EventArgs) Handles TxtTeste1.Click

		Dim Ponto As Point
		Dim RPA As New ClRpa

		Try

			Ponto = RPA.GetMouse()

			'MsgBox(Ponto.X & "|" & Ponto.Y)


		Catch ex As Exception

		End Try

	End Sub
End Class
