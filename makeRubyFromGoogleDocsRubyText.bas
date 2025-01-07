Option Explicit

Public Sub ルビふり()
    Call SetRubyText(ActiveDocument)
End Sub

Private Sub SetRubyText(Optional ByVal a_Document As Document)
    Dim tgtDoc As Document
    Set tgtDoc = a_Document
    Dim tgtParagraph As Paragraph
    Dim tgtRange As Range
    Dim tgtStr As String
    Dim tgtRuby As String
    Dim regPattern As Object
    Dim matches As Object
    Dim match As Object
    Dim i As Long
    
    '正規表現オブジェクト生成
    Set regPattern = New RegExp
    regPattern.Global = True
    regPattern.IgnoreCase = False
    regPattern.Pattern = "\|(.+?)\|\{(.+?)\}"
    
    'パラグラフごとにルビ振り処理
    For Each tgtParagraph In tgtDoc.Paragraphs
        Set matches = regPattern.Execute(tgtParagraph.Range.Text)
        
        '正規表現のパターンに一致した箇所にルビを段落の後ろからふる
        For i = matches.Count - 1 To 0 Step -1
            Set match = matches(i)
            tgtStr = match.Submatches(0)            '対象
            tgtRuby = match.Submatches(1)          'ルビ
            
            '段落内の該当テキストの位置を特定する
            With tgtParagraph.Range
                Call .SetRange(.Start + match.FirstIndex + 1, .Start + match.FirstIndex + Len(tgtStr) + 1)
                Call .PhoneticGuide(Text:=tgtRuby, Alignment:=wdPhoneticGuideAlignmentCenter, Raise:=0, FontSize:=7)
            End With
        Next i
    Next
    
    Call 不要な記号の削除
    
End Sub

Private Sub 不要な記号の削除()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\{*\}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "|"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


