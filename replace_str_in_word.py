from pathlib import Path  # core python module

import win32com.client  # pip install pywin32

# Path settings
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "input"
output_dir = current_dir / "output"
output_dir.mkdir(parents=True, exist_ok=True)

# Find & replace
find_str = "2022"
replace_with = "2023"
wd_replace = 2  # 2=replace all occurences, 1=replace one occurence, 0=replace no occurences
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached

# Open Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False


for doc_file in Path(input_dir).rglob("*.doc*"):
    # Open each document and replace string
    word_app.Documents.Open(str(doc_file))
    # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
    word_app.Selection.Find.Execute(
        FindText=find_str,
        ReplaceWith=replace_with,
        Replace=wd_replace,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )

    # -- Replace str in shapes
    # VBA SO reference: https://stackoverflow.com/a/26266598
    # Loop through all the shapes
    for i in range(word_app.ActiveDocument.Shapes.Count):
        if word_app.ActiveDocument.Shapes(i + 1).TextFrame.HasText:
            words = word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words
            # Loop through each word. This method preserves formatting.
            for j in range(words.Count):
                # If a word exists, replace the text of it, but keep the formatting.
                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with

    # Save the new file
    output_path = output_dir / f"{doc_file.stem}_replaced{doc_file.suffix}"
    word_app.ActiveDocument.SaveAs(str(output_path))
    word_app.ActiveDocument.Close(SaveChanges=False)
word_app.Application.Quit()