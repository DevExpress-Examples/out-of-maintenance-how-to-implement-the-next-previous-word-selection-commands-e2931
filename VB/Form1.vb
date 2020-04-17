Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit
Imports DevExpress.Services
Imports DevExpress.XtraRichEdit.Services.Implementation
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraRichEdit.Internal
Imports DevExpress.Portable.Input

Namespace RichWordsIterator
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			LoadContent()
			SubstituteKeyboardService()
		End Sub

		Private Sub LoadContent()
			richEditControl1.LoadDocument(Application.StartupPath & "\..\..\" & "TextFile.txt", DocumentFormat.PlainText)
		End Sub

		Private Sub SubstituteKeyboardService()
			Dim service As IKeyboardHandlerService = richEditControl1.GetService(Of IKeyboardHandlerService)()
			Dim wrapper As New MyKeyboardHandlerServiceWrapper(service)
			wrapper.RichEditControl = richEditControl1
			richEditControl1.RemoveService(GetType(IKeyboardHandlerService))
			richEditControl1.AddService(GetType(IKeyboardHandlerService), wrapper)
		End Sub
	End Class

	Public Class MyKeyboardHandlerServiceWrapper
		Inherits KeyboardHandlerServiceWrapper
		Public Sub New(ByVal service As IKeyboardHandlerService)
			MyBase.New(service)
		End Sub

		Private privateRichEditControl As RichEditControl
		Public Property RichEditControl() As RichEditControl
			Get
				Return privateRichEditControl
			End Get
			Set(ByVal value As RichEditControl)
				privateRichEditControl = value
			End Set
		End Property

		Public Overrides Sub OnKeyDown(ByVal e As PortableKeyEventArgs)
			Select Case e.KeyCode
				Case PortableKeys.F6
					SelectNextWord()
				Case PortableKeys.F7
					SelectPreviousWord()
				Case Else
					MyBase.OnKeyDown(e)
			End Select
		End Sub

		Private Sub SelectNextWord()
			Dim targetControl As RichEditControl = RichEditControl

			If targetControl.Document.Selection.End.ToInt() = targetControl.Document.Range.End.ToInt() - 1 Then
				Return
			End If

			targetControl.Document.CaretPosition = targetControl.Document.CreatePosition(targetControl.Document.Selection.Start.ToInt() + 1)

			Dim nextWordCommand As New NextWordCommand(targetControl)
			nextWordCommand.Execute()

			Dim extendNextWordCommand As New ExtendNextWordCommand(targetControl)
			extendNextWordCommand.Execute()
		End Sub

		Private Sub SelectPreviousWord()
			Dim targetControl As RichEditControl = RichEditControl

			If targetControl.Document.Selection.Start.ToInt() = 0 Then
				Return
			End If

			targetControl.Document.CaretPosition = targetControl.Document.CreatePosition(targetControl.Document.Selection.Start.ToInt() - 1)

			Dim previousWordCommand As New PreviousWordCommand(targetControl)
			previousWordCommand.Execute()

			Dim extendNextWordCommand As New ExtendNextWordCommand(targetControl)
			extendNextWordCommand.Execute()
		End Sub

	End Class

End Namespace