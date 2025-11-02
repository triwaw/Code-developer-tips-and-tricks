Attribute VB_Name = "Module2"
Sub RemoveFigureTags()

    ' Macro to find and remove all content between (and including)
    ' the opening tag "<figure class="embedded-preview">" and the closing tag "</figure>".
    ' It uses the Find/Replace function with Wildcards enabled.

    ' The search pattern is now: <figure class="embedded-preview">*</figure>
    ' The wildcard * matches any sequence of zero or more characters,
    ' making it the most robust choice for multi-line deletion.

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
        ' The pattern is constructed to find the literal opening tag, followed by
        ' any characters (*), followed by the literal closing tag.
        ' Angle brackets (< and >) must be escaped with a backslash (\) when using wildcards.
        ' Chr(34) is used to correctly represent the double quote character (") within the string.
        .Text = "\<figure class=" & Chr(34) & "embedded-preview" & Chr(34) & "\>*\</figure\>"

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
    MsgBox "All contents within the specific <figure> tags (including the tags themselves) have been removed.", vbInformation

End Sub

