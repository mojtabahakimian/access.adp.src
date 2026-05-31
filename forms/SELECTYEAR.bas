Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8050
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =5520
    Top =2370
    Right =13575
    Bottom =6090
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x682fa12d8a07e340
    End
    Caption ="انتخاب سال مالي پيشرفته"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3741
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =226
                    Top =170
                    Width =7605
                    Height =3288
                    FontSize =14
                    FontWeight =700
                    Name ="DBLIST"
                    RowSourceType ="Value List"
                    ColumnWidths ="1440;1440;1440"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Traffic"
                    OnKeyPress ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "SELECTYEAR.cls"
