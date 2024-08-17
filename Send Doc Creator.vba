' Send Doc Creator v1.0.1---Fully Automated Edition
' https://github.com/KSXia/Verbatim-Send-Doc-Creator/tree/Fully-Automated-Edition
' Updated on 2024-08-17
' Thanks to Truf for providing the original macro this macro is based on!
' This macro has limited compatibility: it may not work with documents saved on MacOS.
Sub CreateAndSaveSendDoc()
	Dim StylesToDelete() As Variant
	
	Dim AutomaticallyCloseSendDoc As Boolean
	
	' ---USER CUSTOMIZATION---
	' <<SET THE STYLES TO DELETE HERE!>>
	' Add the names of styles that you want to delete to this list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas.
	' If the list is empty, the macro will still work, but no styles will be deleted.
	StylesToDelete = Array("Analytic", "Undertag")
	
	' <<SET WHETHER TO AUTOMATICALLY CLOSE THE SEND DOC AFTER IT'S CREATED AND SAVED HERE!>>
	' If AutomaticallyCloseSendDoc is set to True, the send doc will be automatically closed after it is created and saved.
	AutomaticallyCloseSendDoc = True
	
	' ---INITIAL VARIABLE SETUP---
	Dim OriginalDoc As Document
	' Assign the original document to a variable
	Set OriginalDoc = ActiveDocument
	
	' Check if the original document has previously been saved
	If OriginalDoc.Path = "" Then
		' If the original document has not been previously saved:
		MsgBox "The current document must be saved at least once. Please save the current document and try again.", Title:="Error in Creating Send Doc"
		Exit Sub
	End If
	
	' Assign the original document name to a variable
	Dim OriginalDocName As String
	OriginalDocName = OriginalDoc.Name
	
	Dim SendDoc As Document
	
	' If the doc has been previously saved, create a copy of it to be the send doc
	Set SendDoc = Documents.Add(OriginalDoc.FullName)
	
	Dim GreatestStyleIndex As Integer
	GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
	
	' ---INITIAL GENERAL SETUP---
	' Disable error prompts in case one of the styles set to be deleted isn't present
	On Error Resume Next
	
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---STYLE DELETION---
	Dim CurrentStyleIndex As Integer
	For CurrentStyleIndex = 0 to GreatestStyleIndex Step +1
		Dim StyleToDelete As Style
		
		' Specify the style to be deleted and delete it
		Set StyleToDelete = SendDoc.Styles(StylesToDelete(CurrentStyleIndex))
		
		' Use Find and Replace to remove text with the specified style and delete it
		With SendDoc.Content.Find
			.ClearFormatting
			.Style = StyleToDelete
			.Replacement.ClearFormatting
			.Replacement.Text = ""
			.Format = True
			' Disabling checks in the find process for optimization
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			' Delete all text with the style to delete
			.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
		End With
	Next CurrentStyleIndex
	
	' ---POST STYLE DELETION PROCESSES---
	' Re-enable error prompts
	On Error GoTo 0
	
	' ---SAVING THE SEND DOC---
	Dim SavePath As String
	SavePath = OriginalDoc.Path & "\" & Left(OriginalDocName, Len(OriginalDocName) - 5) & " [S]" & ".docx"
	SendDoc.SaveAs2 Filename:=SavePath, FileFormat:=wdFormatDocumentDefault
	
	If AutomaticallyCloseSendDoc Then
		SendDoc.Close SaveChanges:=wdSaveChanges
		MsgBox "The send doc is saved at " & SavePath, Title="Successfully Created and Saved Send Doc"
	End If
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub
