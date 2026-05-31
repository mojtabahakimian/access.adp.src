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
    Width =8390
    DatasheetFontHeight =10
    ItemSuffix =25
    Left =5985
    Top =2565
    Right =14370
    Bottom =5505
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="."
    OnOpen ="[Event Procedure]"
    HelpFile ="amin98.hlp>main"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
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
            Height =2948
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3872
                    Top =2499
                    Width =870
                    Height =315
                    FontWeight =700
                    TabIndex =2
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3004
                    Top =2499
                    Width =870
                    Height =317
                    FontWeight =700
                    TabIndex =3
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7205
                    Top =623
                    Width =729
                    Height =255
                    TabIndex =4
                    Name ="DTT"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2485
                    Top =118
                    Width =3396
                    Height =555
                    FontSize =14
                    FontWeight =700
                    BackColor =13434828
                    Name ="SNDNUM1"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    FontName ="Titr"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5967
                            Top =218
                            Width =1125
                            Height =285
                            Name ="Label17"
                            Caption ="قيمت  پايه:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =1530
                    Top =1814
                    Width =5235
                    Height =615
                    FontSize =10
                    FontWeight =700
                    ForeColor =128
                    Name ="Label20"
                    Caption ="قيمت    =       مبلغ پايه     +  مبلغ پايه   *   درصد     +   مبلغ بالانس"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2490
                    Top =735
                    Width =3396
                    Height =555
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    BackColor =10092543
                    Name ="DIG"
                    DefaultValue ="2"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =5975
                            Top =842
                            Width =1500
                            Height =285
                            Name ="Label24"
                            Caption ="تعداد ارقام رند شدن:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ZCHANGE_FI.cls"
