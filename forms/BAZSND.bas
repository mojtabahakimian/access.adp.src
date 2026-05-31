Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5839
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =3030
    Top =2550
    Right =8865
    Bottom =4185
    HelpContextId =131
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbece4ccb101ce240
    End
    Caption ="بازسازي  اسناد"
    HelpFile ="AMIN98.HLP>MAIN"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =1644
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextFontFamily =10
                    BackStyle =0
                    Left =457
                    Top =765
                    Width =4743
                    Height =180
                    FontSize =6
                    ForeColor =8388608
                    Name ="Text0"
                    DefaultValue ="\"n\""
                    FontName ="Wingdings"
                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =450
                    Top =766
                    Width =4743
                    Height =185
                    Name ="Box7"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =42
                    Left =2205
                    Top =1170
                    Width =1245
                    Height =405
                    FontWeight =700
                    TabIndex =1
                    Name ="Command15"
                    Caption ="ادامه عمليات"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =42
                    Left =172
                    Top =113
                    Width =5385
                    Height =615
                    FontSize =10
                    FontWeight =700
                    Name ="Label16"
                    Caption ="مطمئن شويد كه فرد ديگري از سيستم استفاده نمي كند تا مرتب سازي اسناد صورت گيرد..."
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "BAZSND.cls"
