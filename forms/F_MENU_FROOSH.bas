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
    Width =5783
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =5700
    Top =3255
    Right =11490
    Bottom =5280
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="گزارش گروهي فاكتورها"
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
        Begin OptionButton
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
            Height =2040
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2520
                    Top =1417
                    Width =900
                    Height =397
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
                    Left =1644
                    Top =1417
                    Width =870
                    Height =397
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
                    Left =4280
                    Top =623
                    Width =729
                    Height =255
                    TabIndex =4
                    Name ="DTT"
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2880
                    Top =180
                    Width =1431
                    Height =300
                    Name ="SNDNUM1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT NUMBER FROM HEAD_LST GROUP BY NUMBER, TAG HAVING (TAG = 2"
                        ") ORDER BY NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4395
                            Top =180
                            Width =1140
                            Height =285
                            Name ="Label17"
                            Caption ="از  شماره:"
                        End
                    End
                End
                Begin ComboBox
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =225
                    Top =170
                    Width =1431
                    Height =300
                    TabIndex =1
                    Name ="SNDNUM2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT NUMBER FROM HEAD_LST GROUP BY NUMBER, TAG HAVING (TAG = 2"
                        ") ORDER BY NUMBER DESC"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1700
                            Top =165
                            Width =1125
                            Height =285
                            Name ="Label4"
                            Caption ="تا  شماره:"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =247
                    Left =4422
                    Top =566
                    Width =1247
                    Height =1273
                    TabIndex =5
                    Name ="KINDP"
                    DefaultValue ="2"
                    Begin
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5223
                            Top =804
                            OptionValue =1
                            Name ="Option23"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =4608
                                    Top =776
                                    Width =615
                                    Height =240
                                    Name ="Label24"
                                    Caption ="عمومي"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5223
                            Top =1134
                            OptionValue =2
                            Name ="Option25"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =4878
                                    Top =1106
                                    Width =345
                                    Height =240
                                    Name ="Label26"
                                    Caption ="ويژه"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5223
                            Top =1464
                            OptionValue =3
                            Name ="Option27"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =4788
                                    Top =1436
                                    Width =435
                                    Height =240
                                    Name ="Label28"
                                    Caption ="حواله"
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_FROOSH.cls"
