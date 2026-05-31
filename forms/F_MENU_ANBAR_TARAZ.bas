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
    Width =5952
    DatasheetFontHeight =10
    ItemSuffix =34
    Left =6570
    Top =3075
    Right =12525
    Bottom =4710
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="تراز انبار"
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
            Height =1644
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2893
                    Top =1125
                    Width =870
                    Height =375
                    FontWeight =700
                    TabIndex =3
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2025
                    Top =1125
                    Width =870
                    Height =377
                    FontWeight =700
                    TabIndex =4
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =687
                    Top =645
                    Width =234
                    Height =255
                    TabIndex =7
                    Name ="DTT"
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
                    Left =3334
                    Top =225
                    Width =1431
                    Height =300
                    Name ="CANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE FROM TCOD_ANBAR GROUP BY TCOD_ANBAR.CODE HAVING (((TCOD_A"
                        "NBAR.CODE)<>0));"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4859
                            Top =227
                            Width =495
                            Height =285
                            Name ="Label4"
                            Caption ="انبار:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =390
                    Top =239
                    Width =2841
                    Height =300
                    TabIndex =1
                    Name ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES FROM TCOD_ANBAR GROUP BY TCOD_ANBAR.COD"
                        "E, TCOD_ANBAR.NAMES HAVING (((TCOD_ANBAR.CODE)<>0)) ORDER BY TCOD_ANBAR.NAMES;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =390
                    Top =645
                    Width =234
                    Height =255
                    TabIndex =5
                    Name ="MANBAR"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =390
                    Top =1020
                    Width =234
                    Height =255
                    TabIndex =6
                    Name ="MKALA"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3344
                    Top =623
                    Width =1416
                    Height =300
                    TabIndex =2
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4854
                            Top =625
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_ANBAR_TARAZ.cls"
