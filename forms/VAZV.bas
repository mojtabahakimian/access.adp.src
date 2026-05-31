Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
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
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =1700
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =8700
    Top =2730
    Right =10395
    Bottom =5370
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x00d79ef8d357e440
    End
    Caption ="انتخاب وضعيت"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
            Height =2654
            Name ="Detail"
            Begin
                Begin ListBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    OldBorderStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Width =1695
                    Height =2654
                    FontSize =10
                    FontWeight =700
                    Name ="List0"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\";7;\"حذف شده\""
                    ColumnWidths ="567;567"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Titr"
                    OnKeyPress ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "VAZV.cls"
