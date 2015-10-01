 With ActiveDocument.Styles("Title").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(-1)
        .RightIndent = CentimetersToPoints(0)
        .TabStops.ClearAll
       .TabStops.Add Position:=CentimetersToPoints(0), Alignment:=wdAlignTabLeft, Leader:= _
        wdTabLeaderSpaces

        
    End With
