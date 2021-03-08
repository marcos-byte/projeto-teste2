
#Region "Bibliotecas"

'-----------------------------
'Import library RPA (Web)
'-----------------------------
Imports System
Imports System.Windows.Forms
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Support.UI.WebDriverWait
Imports OpenQA.Selenium.Interactions.Actions
Imports System.Runtime.Remoting.Contexts
Imports System.Runtime.InteropServices

'-----------------------------
'Import library RPA (Desktop)
'-----------------------------
'Imports System
Imports System.Windows.Automation
Imports System.Windows.Automation.ControlType
Imports System.Diagnostics
Imports System.Threading

Imports TestStack.White.Factory
Imports TestStack.White.UIItems
Imports TestStack.White.UIItems.WindowStripControls

Imports WindowsInput.Native
Imports WindowsInput



#End Region

Public Class ClRpa


#Region "Variáveis"

	'-----------------------------
	''Define variáveis Web
	'-----------------------------
	Dim _Url As String
	'Dim webDriver As OpenQA.Selenium.Firefox.FirefoxDriver
	'Dim Act As Interactions.Actions
	'Dim Act As New Interactions.Actions(webDriver)

	'-----------------------------
	''Define variáveis Desktop
	'-----------------------------
	Dim AplDesk As Process
	'Dim WinApl As AutomationElement = AutomationElement.RootElement
	'Dim LstEl As AutomationElementCollection

	Const MOUSEEVENTF_LEFTDOWN As Integer = 2
	Const MOUSEEVENTF_LEFTUP As Integer = 4
	Const MOUSEEVENTF_RIGHTDOWN As Integer = 6
	Const MOUSEEVENTF_RIGHTUP As Integer = 8


	Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As IntPtr)


#End Region


#Region "Construtor"

	'Constructor
	Public Sub New()

	End Sub

#End Region


#Region "Propiedades"

	'-----------------------------
	'Propiedades
	'-----------------------------

	'Url
	Public Property Url() As String

		Get
			Return _Url

		End Get

		Set(ByVal Value As String)

			_Url = Value
		End Set

	End Property

	'Janela
	Public Property Janela() As Process

		Get
			Return AplDesk

		End Get

		Set(ByVal Value As Process)

			AplDesk = Value

		End Set

	End Property


#End Region


#Region "Métodos"

	'---------------------------------------------------------
	'Métodos principais
	'---------------------------------------------------------

	'-------------------------------
	'Métodos referentes ao navegador
	'-------------------------------


	'-------------------------------
	'Métodos referentes ao Mouse
	'-------------------------------


	'-------------------------------
	'Métodos referentes ao teclado
	'-------------------------------




	'---------------------------------------------------------
	'Métodos genéricos
	'---------------------------------------------------------

	'Pause
	Public Function Pause(ByVal RefP As Integer) As Boolean

		Try

			'Pausa
			System.Threading.Thread.Sleep(RefP * 1000)

			'Retorna valor
			Return True

		Catch ex As Exception

			'Retorna valor
			Return False

		End Try

	End Function



	'--------------------------------
	'Métodos específicos para Desktop
	'--------------------------------

	'Executa aplicação 'Desktop'
	Public Function ExecApp(ByVal RefAplDesk As String, ByVal RefPath As String) As Boolean

		'Declara variáveis
		Dim VAplD As String

		Try

			'Define aplicação que deverá ser executada
			VAplD = RefPath & "\" & RefAplDesk

			'Executa aplicação
			AplDesk = Process.Start(VAplD)

			'Retorna valor
			Return True

		Catch ex As Exception

			MsgBox("Erro no ExecApp:" & ex.Message)
			MsgBox("Erro no ExecApp:" & ex.StackTrace)


			'Retorna valor
			Return False


		End Try


	End Function

	'Verifica posição do cursor 
	Public Function GetMouse() As Point

		'Declara variáveis
		Dim Pos As Point

		Try

			'Verifica posição do mouse
			Pos = Cursor.Position

			'Retorna valor
			Return Pos

		Catch ex As Exception

			'Retorna valor
			Return Nothing

		End Try


	End Function

	'Move cursor do mouse para determinada posição na tela 
	Public Function MovMouse(ByVal Refx As Integer, ByVal Refy As Integer) As Boolean

		Try

			'Posiciona cursor do mouse
			Windows.Forms.Cursor.Position = New Point(Refx, Refy)

			'Retorna valor
			Return True

		Catch ex As Exception

			'Retorna valor
			Return False

		End Try


	End Function

	'Click mouse 
	Public Function MouseClick(ByVal RefTipB As String, ByVal RefTipCli As String, ByVal Refx As Integer, ByVal Refy As Integer) As Boolean

		'------------------------------------------  
		'Declara constantes 
		'------------------------------------------ 
		Dim TipB As String
		Dim TipCli As String


		Try

			'--------------------------------------
			'Move cursor para posição específica
			'--------------------------------------

			'Valida Coordenadas do cursor
			If Refx <> 0 And Refy <> 0 Then

				Cursor.Position = New Point(Refx, Refy)

			End If

			'Tipo de botão
			Select Case LCase(RefTipB)
				Case Is = "left", "right"

					TipB = "left"

			End Select

			'Tipo de Click
			Select Case LCase(RefTipCli)
				Case Is = "single", "double"

					TipCli = "single"

			End Select


			'Verifica Duplo click
			Select Case LCase(RefTipB)
				Case Is = "left"

					Select Case LCase(RefTipB)
						Case Is = "single"

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)

							'Solta botão do mouse
							mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)

						Case Is = "double"

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)

							'Solta botão do mouse
							mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, IntPtr.Zero)

							'Solta botão do mouse
							mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, IntPtr.Zero)

					End Select

				Case Is = "right"

					Select Case LCase(RefTipB)
						Case Is = "single"


							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero)

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero)

						Case Is = "double"

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero)

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero)

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, IntPtr.Zero)

							'Pressiona botão do mouse 
							mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, IntPtr.Zero)

						Case Else
					End Select

				Case Else

			End Select

			'Retorna valor
			Return True

		Catch ex As Exception

			'Retorna valor
			Return False

		End Try


	End Function

	'Write keyboard
	Public Function KeyBoard(ByVal RefText As String) As Boolean

		Try

			'Write 
			My.Computer.Keyboard.SendKeys(RefText)

			'Retorna valor
			Return True

		Catch ex As Exception

			'Retorna valor
			Return False

		End Try

	End Function











#End Region

End Class
