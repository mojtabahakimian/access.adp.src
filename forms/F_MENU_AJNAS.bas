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
    Width =6292
    DatasheetFontHeight =10
    ItemSuffix =38
    Left =4485
    Top =2295
    Right =10770
    Bottom =5745
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="چاپ كالاها"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            FontName ="Tahoma"
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =3458
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3149
                    Top =3106
                    Width =870
                    Height =315
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2281
                    Top =3106
                    Width =870
                    Height =317
                    TabIndex =1
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =1203
                    Top =45
                    Width =3860
                    Height =1287
                    TabIndex =2
                    Name ="CHAP"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =4687
                            Top =279
                            OptionValue =1
                            Name ="Option23"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =2782
                                    Top =253
                                    Width =1905
                                    Height =255
                                    Name ="Label24"
                                    Caption ="قيمت عمده چاپ شود"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =4672
                            Top =624
                            OptionValue =3
                            Name ="Option19"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =2767
                                    Top =598
                                    Width =1905
                                    Height =255
                                    Name ="Label20"
                                    Caption ="قيمت خرده چاپ شود"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =4672
                            Top =954
                            OptionValue =2
                            Name ="Option21"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =2707
                                    Top =926
                                    Width =1965
                                    Height =240
                                    Name ="Label22"
                                    Caption ="قيمت كالا چاپ نشود"
                                End
                            End
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =4622
                    Top =1444
                    TabIndex =3
                    Name ="M0"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1664
                            Top =1410
                            Width =2955
                            Height =240
                            Name ="Label33"
                            Caption ="كالاها باموجودي صفر نمايش داده شوند"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =4622
                    Top =1759
                    TabIndex =4
                    Name ="cdd"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1664
                            Top =1725
                            Width =2955
                            Height =240
                            Name ="Label35"
                            Caption ="كد كالا چاپ شود"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =240
                    Top =2199
                    Width =5782
                    Height =623
                    TabIndex =5
                    Name ="MATN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4901
                            Top =1915
                            Width =1065
                            Height =240
                            Name ="Label37"
                            Caption ="متن زير گزارش"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_AJNAS.cls"
