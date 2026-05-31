Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5555
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =3180
    Top =2115
    Right =8730
    Bottom =4935
    HelpContextId =131
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbece4ccb101ce240
    End
    Caption ="دريافت اطلاعات"
    OnCurrent ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
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
            Height =2834
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =207
                    TextFontCharSet =0
                    Left =2040
                    Top =2154
                    Width =1644
                    Height =510
                    FontWeight =700
                    Name ="Command2"
                    Caption ="&دريافت اطلاعات"
                    FontName ="Tahoma"
                    ControlTipText ="اين گزينه را جهت دريافت  اطلاعات از ديسك  تاييد نماييد"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =241
                    Width =2896
                    Height =300
                    TabIndex =1
                    Name ="Text7"
                    DefaultValue ="\"A:\\\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =3560
                            Top =226
                            Width =1485
                            Height =285
                            Name ="Label8"
                            Caption ="مسير فايل اطلاعات :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =420
                    Top =900
                    Width =4740
                    Height =855
                    Name ="Label9"
                    Caption ="نام درايو و مسيري كه اطلاعات بايد از آنجا خوانده شود را بطور كامل وارد نماييد. ن"
                        "ام فايل لازم نيست نوشته شود. سپس گزينه دريافت اطلاعات را كليك كنيد"
                    FontName ="Titr"
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =453
                    Top =113
                    Width =4706
                    Height =624
                    Name ="Box11"
                End
            End
        End
    End
End
CodeBehindForm
' See "F_DARYAFT.cls"
