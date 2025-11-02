Attribute VB_Name = "Module1"
Sub RemoveFigureTagsAndContent()
    ' Macro to find and remove all content between (and including)
    ' the opening tag "<figure" and the closing tag "</figure>".
    ' It uses the Find/Replace function with Wildcards enabled.

    ' The search pattern is: <figure*</figure>

    Dim doc As Document
    Set doc = ActiveDocument

    ' Check if there is an active document
    If doc Is Nothing Then
        MsgBox "No active document found. Please open a Word document and try again.", vbExclamation
        Exit Sub
    End If

    With doc.Content.Find
        ' Reset Find/Replace settings
        .ClearFormatting
        .Replacement.ClearFormatting

        ' 1. Set the search string (Find What)
        ' <figure: finds the literal opening tag
        ' *: finds any character zero or more times
        ' </figure>: finds the literal closing tag
        ' IMPORTANT: Word's wildcard search is non-greedy by default
        ' if using the * as shown, but it is often safer to use the
        ' technique below for multi-line matches.
        ' However, for simple removal, this often works best.
        .Text = "\<figure*\</figure>"

        ' 2. Set the replacement string (Replace With)
        ' We replace the matched text with an empty string ("") to delete it
        .Replacement.Text = ""

        ' 3. Configure the search options
        .Forward = True ' Search from start to end
        .Wrap = wdFindContinue ' Continue searching from the start after reaching the end
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchFuzzy = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = False

        ' CRITICAL: Enable Wildcards for the pattern to work
        .MatchWildcards = True

        ' 4. Execute the replace operation (Replace All)
        ' wdReplaceAll executes the replacement for all found instances
        .Execute Replace:=wdReplaceAll

    End With

    ' Inform the user about the result
    MsgBox "All contents within the <figure> tags (including the tags themselves) have been removed.", vbInformation

End Sub

