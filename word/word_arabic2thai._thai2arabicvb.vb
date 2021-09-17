Option Explicit

Sub arabic2thai()
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Dim i As Integer
    For i = 0 To 9
        'Debug.Print Chr(i)
        'Debug.Print ChrW(3664 - 48 + i)
        With Selection.Find
            .Text = ChrW(48 + i) ' find 0-9 arabic numerals
            .Replacement.Text = ChrW(3664 + i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

Sub thai2arabic()
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Dim i As Integer
    For i = 0 To 9 
        With Selection.Find
            .Text = ChrW(3664 + i) ' find 0-9 Thai numerals
            .Replacement.Text = ChrW(48 + i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

