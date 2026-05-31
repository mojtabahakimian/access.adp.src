Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
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
    Width =6645
    DatasheetFontHeight =10
    ItemSuffix =34
    Left =2940
    Top =2190
    Right =8970
    Bottom =4050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x41a653f29e1ae240
    End
    Caption =" "
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin Line
            Width =1701
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
            FontName ="MS Sans Serif"
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
            Height =3571
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =2820
                    Top =3015
                    Width =1281
                    Height =418
                    FontWeight =700
                    Name ="Command1"
                    Caption ="&تائيد"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =26
                    Top =56
                    Width =915
                    Height =1020
                    FontSize =48
                    Name ="Label13"
                    Caption ="V"
                    FontName ="Webdings"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =223
                    Left =850
                    Top =221
                    Width =170
                    Height =227
                    Name ="Line14"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =215
                    Left =907
                    Top =448
                    Width =226
                    Height =113
                    Name ="Line15"
                End
                Begin Line
                    OverlapFlags =87
                    Left =915
                    Top =670
                    Width =283
                    Height =57
                    Name ="Line16"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =225
                    TextFontCharSet =0
                    Left =1530
                    Top =3015
                    Width =1281
                    Height =418
                    FontWeight =700
                    TabIndex =1
                    Name ="Command17"
                    Caption ="&لغو"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =1757
                    Top =56
                    Width =3645
                    Height =345
                    Name ="Label18"
                    Caption ="اعلام وصول چكهاي انتخاب شده  انجام شود؟"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =120
                    Top =2437
                    Width =5106
                    Height =300
                    TabIndex =2
                    Name ="HHMOIN"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;3686;567"
                    DefaultValue ="1"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5265
                            Top =2437
                            Width =1275
                            Height =300
                            Name ="Label4"
                            Caption ="وصول به حساب :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3859
                    Top =1814
                    Width =1256
                    Height =315
                    TabIndex =3
                    Name ="DTS"
                    Format ="0000\\/00\\/00"
                    ValidationRule ="Not Is Null"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =5162
                            Top =1818
                            Width =1035
                            Height =300
                            Name ="Label20"
                            Caption ="تاريخ وصول:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =2655
                    Top =566
                    Width =3767
                    Height =1693
                    TabIndex =4
                    Name ="VCLC"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =5054
                            Top =450
                            Width =825
                            Height =240
                            BackColor =-2147483633
                            Name ="Label22"
                            Caption ="تاريخ وصول"
                            FontName ="Tahoma"
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5976
                            Top =804
                            OptionValue =1
                            Name ="Option24"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =3966
                                    Top =776
                                    Width =2010
                                    Height =240
                                    Name ="Label25"
                                    Caption ="درتاريخ سررسيد وصول شود"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5976
                            Top =1134
                            OptionValue =2
                            Name ="Option26"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =2841
                                    Top =1106
                                    Width =2790
                                    Height =240
                                    Name ="Label27"
                                    Caption ="روز پس از تاريخ سررسيد وصول شود"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            ReadingOrder =1
                            Left =5976
                            Top =1464
                            OptionValue =3
                            Name ="Option28"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =4281
                                    Top =1436
                                    Width =1695
                                    Height =240
                                    Name ="Label29"
                                    Caption ="درتاريخ ذيل وصول شود"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5624
                    Top =1110
                    Width =326
                    Height =255
                    FontWeight =700
                    TabIndex =5
                    Name ="DTC"
                    ValidationRule ="Not Is Null"
                    DefaultValue ="1"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "VOSULDIALOG2.cls"
