' ---Send Doc Creator v1.1.1---
' Updated on 2024-08-22.
' Basic Edition: This edition of the Send Doc Creator only has the style deleting mechanism and does not have any mechanisms regarding saving the send doc.
' https://github.com/KSXia/Verbatim-Send-Doc-Creator
' Thanks to Truf for providing the original macro this macro is based on!
Sub CreateSendDoc()
	Dim DeleteStyles As Boolean
	Dim StylesToDelete() As Variant
	
	' ---USER CUSTOMIZATION---
	' <<SET THE STYLES TO DELETE HERE!>>
	' Add the names of styles that you want to delete to this list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas.
	' If the list is empty, the macro will still work, but no styles will be deleted.
	StylesToDelete = Array("Undertag", "Analytic")
	
	' If DeleteStyles is set to True, the styles listed in the StylesToDelete array will be deleted. If DeleteStyles is set to False, these styles will not be deleted.
	' If you want to disable the deletion of these styles, set DeleteStyles to False.
	DeleteStyles = True
	
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
	
	' ---INITIAL GENERAL SETUP---
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---VARIABLE SETUP---
	Dim SendDoc As Document
	
	' If the doc has been previously saved, create a copy of it to be the send doc
	Set SendDoc = Documents.Add(OriginalDoc.FullName)
	
	Dim GreatestStyleIndex As Integer
	GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
	
	' ---STYLE DELETION SETUP---
	' Disable error prompts in case one of the styles set to be deleted isn't present
	On Error Resume Next
	
	' ---STYLE DELETION---
	If DeleteStyles = True Then
		Dim CurrentStyleToDeleteIndex As Integer
		For CurrentStyleToDeleteIndex = 0 to GreatestStyleIndex Step 1
			Dim StyleToDelete As Style
			
			' Specify the style to be deleted
			Set StyleToDelete = SendDoc.Styles(StylesToDelete(CurrentStyleToDeleteIndex))
			
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
		Next CurrentStyleToDeleteIndex
	End If
	
	' ---POST STYLE DELETION PROCESSES---
	' Re-enable error prompts
	On Error GoTo 0
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub
