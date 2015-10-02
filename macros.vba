Sub setup()
  Call outdentheadings
   DefineIndentStyles
 Call markit
End Sub

Sub markit()
 ActiveWindow.View.ShowHiddenText = True
  With Options
        .AutoFormatAsYouTypeApplyBulletedLists = False
        
        '.ShowHiddenText = True
    End With
    
   'Call labelImages
    Dim doc As Document
    Dim para As Paragraph
    
    'Headings
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "   "
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Dim tabb As String
    tabb = Chr(9)

    Set doc = ActiveDocument
    For Each para In Selection.Paragraphs
        Dim contents As String
    contents = para.Range.Text
    numtabs = UBound(Split(contents, tabb))
    If numtabs < 6 Then
        If InStr(contents, "title:") = 1 Then
             para.Style = "Title"
             
           para.Range.Find.ClearFormatting
            para.Range.Find.Replacement.ClearFormatting
            With Selection.Find
                    .Text = "title:"
                    .Replacement.Text = ""
                     .Forward = True
                      .Wrap = wdFindContinue
                      .Format = True
                  .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Font.Italic = False
            With .Replacement
                .Font.Hidden = True
                .Font.Size = "12"
            End With
            End With
             Selection.Find.Execute Replace:=wdReplaceAll
       ElseIf InStr(contents, "####") = 1 Then
             para.Style = "Heading 4"
       ElseIf InStr(contents, "###") = 1 Then
             para.Style = "Heading 3"
       ElseIf InStr(contents, "##") = 1 Then
             para.Style = "Heading 2"
       ElseIf InStr(contents, "#") = 1 Then
             para.Style = "Heading 1"
       ElseIf numtabs > 0 Then
                para.Style = "Indent" + Trim(Str(numtabs))
       Else
             para.Style = "Normal"
       End If
  End If



Next
     
  FormatPara



End Sub




Sub FormatPara()
'
' Macro2 Macro
'
'
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[\*][\*]*[\*][\*]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Font.Italic = False
        With .Replacement
            .Font.Italic = True
            .Font.Bold = True
        End With
        
    End With
    
    'Emph
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[\*]*[\*]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Font.Italic = False
        With .Replacement
            .Font.Italic = True
        End With
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Rem hide formatting cues so the doc will print nicely
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[\*]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Font.Italic = True
        With .Replacement
            .Font.Hidden = True
        End With
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'Headings
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(#{1,})[ ^t]{1,}"
        .Replacement.Text = "\1^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        With .Replacement
            .Font.Hidden = True
        End With
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
      'Headings
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[ ^t]{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        With .Replacement
            .Font.Hidden = False
        End With
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
End Sub
Sub Macro1()
'
' Macro1 Macro
'
'
End Sub
Sub DefineIndentStyles()
'
' Macro3 Macro
'
For num = 1 To 5
    StyleName = "Indent" + CStr(num)
    
    On Error Resume Next
    ActiveDocument.Styles.Add name:=StyleName, Type:=wdStyleTypeParagraph
    With ActiveDocument.Styles(StyleName)
        .AutomaticallyUpdate = False
        .BaseStyle = "Plain Text"
    End With
    
    With ActiveDocument.Styles(StyleName).ParagraphFormat
        .LeftIndent = CentimetersToPoints(num)
        .RightIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(-num)
    End With
    
    ActiveDocument.Styles(StyleName).ParagraphFormat.TabStops.ClearAll
    i = 1
    While i <= num
        ActiveDocument.Styles(StyleName).ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(i), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces
        i = i + 1
    Wend
    Next
    Exit Sub
Leave:
    
End Sub


Sub outdentheadings()
'
' headingss Macro
'
'
   For i = 1 To 5:
    With ActiveDocument.Styles("Heading " & CStr(i)).ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(-1)
        .RightIndent = CentimetersToPoints(0)
        .TabStops.ClearAll
       .TabStops.Add Position:=CentimetersToPoints(0), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces

        
    End With
   
Next

 With ActiveDocument.Styles("Title").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(-1)
        .RightIndent = CentimetersToPoints(0)
        .TabStops.ClearAll
       .TabStops.Add Position:=CentimetersToPoints(0), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces

        
    End With
End Sub

  Sub saveAsDoc(name)
   ActiveDocument.SaveAs fileName:=name, FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, HTMLDisplayOnlyOutput:=False, MaintainCompat:= _
        False
  End Sub
  
 Sub loseTabs()
  Selection.WholeStory
With Selection.Find
    .ClearFormatting
    .Text = "\>^t"
    .Replacement.ClearFormatting
    .Replacement.Text = ">    "
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
   End With
   
With Selection.Find
    .ClearFormatting
    .Text = "-^t"
    .Replacement.ClearFormatting
    .Replacement.Text = "-    "
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
End With

     With Selection.Find
    .ClearFormatting
    .Text = Chr(9)
    .Replacement.ClearFormatting
    .Replacement.Text = "    "
    .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
End With
End Sub

Sub saveit()


 Call markit
 
 Set doc = ActiveDocument
 Dim name As String
 'Remember this doc's name
 name = ActiveDocument.FullName
 fileName = ActiveDocument.name

  Selection.WholeStory
    Selection.Copy
    Documents.Add DocumentType:=wdNewBlankDocument
    Selection.PasteAndFormat (wdPasteDefault)
    Call loseTabs
  ActiveDocument.SaveAs fileName:=name + ".md.html", FileFormat:=wdFormatHTML, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False, HTMLDisplayOnlyOutput:=False, MaintainCompat:=False

    ActiveDocument.SaveAs fileName:=name + ".md", FileFormat:=wdFormatText, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False, Encoding:=65001, InsertLineBreaks:=False, AllowSubstitutions:= _
        False, LineEnding:=wdCRLF

  ActiveWindow.Close
  Dim path As Variant
  
  
  pathArray = Split(name, ":")
  path = "/"
  
  For dirnum = 1 To UBound(pathArray) - 1
    path = path + pathArray(dirnum) + "/"
 Next


End Sub


