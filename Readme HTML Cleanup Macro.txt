🧹 Word Document Cleanup Macro: Figure Tag Remover
A simple, yet powerful Visual Basic for Applications (VBA) macro for Microsoft Word designed to clean up documents by removing specific blocks of HTML-like tags and all their contents.
This macro is particularly useful when converting content from web sources or CMS systems that embed formatting or figure data using recognizable tags.
💡 The Problem Solved
The macro finds and deletes all text, images, and content blocks that start with the specific HTML tag:
<figure class="embedded-preview">
and continue until the closing tag:
</figure>
It removes both the tags and all content enclosed between them, which is essential for quickly normalizing document text.
🛠️ Macro Code: RemoveFigureTagsAndContent
The core logic uses Word's Wildcard search feature for robust, multi-line matching.
How it Works: The Search Pattern
The VBA code uses the following precise wildcard pattern:
.Text = "\<figure class=" & Chr(34) & "embedded-preview" & Chr(34) & "\>*\</figure\>"
• \<\>: Escapes the angle brackets, forcing the search to look for the literal characters < and >.
• Chr(34): Inserts the literal double-quote character ("), ensuring the class name is matched exactly.
• *: The essential wildcard that matches any sequence of zero or more characters, crucially allowing the match to span multiple lines, paragraphs, and sections.
The macro then replaces this entire matched block (Find What) with an empty string (Replace With), effectively deleting it.
🚀 Installation and Usage
Step 1: Open the VBA Editor
1. Open the Word document you want to run the macro on.
2. Press ALT + F11 on your keyboard to open the Visual Basic for Applications (VBA) Editor.
Step 2: Insert a New Module
1. In the VBA Editor, go to the Project Explorer pane (usually on the left).
2. Right-click on your document's project (e.g., Project (Document1) or Normal) and select Insert > Module.
3. Copy the full code from Module1.bas and paste it into the empty code window.
Step 3: Run the Macro
1. Return to your Word document.
2. Go to the View tab on the ribbon.
3. Click Macros > View Macros (or press ALT + F8).
4. Select the macro named RemoveFigureTagsAndContent from the list.
5. Click Run.
A message box will pop up confirming that the process is complete.
🤝 Contribution & License
This project is shared to benefit the community and encourage learning. Feel free to use and adapt this code in your own work.
This macro is released under the MIT License.

