Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =2608
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =360
    Top =4560
    Right =3000
    Bottom =6000
    RecSrcDt = Begin
        0x97a17a1adbc1e540
    End
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SplitFormSplitterBar =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin FormHeader
            Height =453
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Width =2608
                    Height =453
                    Name ="Command2"
                    Caption ="مديريت مالي و بودجه"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =4365
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =2495
                    Height =426
                    Name ="p1e4"
                    Caption ="بودجه"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =627
                    Width =2495
                    Height =426
                    Name ="p2e4"
                    Caption ="ليست بودجه"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =1194
                    Width =2495
                    Height =426
                    Name ="P3E4"
                    Caption ="تحليل بودجه"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "f9.cls"
