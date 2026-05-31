Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5062
    RowHeight =330
    ItemSuffix =6
    Left =5430
    Top =1710
    Right =10965
    Bottom =4980
    HelpContextId =42
    DatasheetGridlinesColor =12632256
    Filter ="((NAMES ALike \"%رفاه%\"))"
    OrderBy ="CODE"
    RecSrcDt = Begin
        0xdc81b0873b9be240
    End
    RecordSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, TCOD_ANBAR.KIND FROM TCOD_ANBAR WHERE "
        "(((TCOD_ANBAR.CODE)<>0));"
    Caption ="تعريف انبارها"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1587
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2718
                    Top =113
                    Width =1416
                    Height =255
                    ColumnWidth =540
                    Name ="CODE"
                    ControlSource ="CODE"
                    ValidationRule =">0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4205
                            Top =113
                            Width =330
                            Height =285
                            Name ="Label1"
                            Caption ="كد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =624
                    Top =453
                    Width =3510
                    Height =255
                    ColumnWidth =3435
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4220
                            Top =453
                            Width =585
                            Height =285
                            Name ="Label3"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1635
                    Top =855
                    Width =2496
                    Height =255
                    ColumnWidth =1995
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="0;'خدمات';1;'مواد مصرفي توليد';2;'نيمه ساخته';3;'ساخته شده';4;'ساير'"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4206
                            Top =860
                            Width =645
                            Height =285
                            Name ="Label5"
                            Caption ="نوع انبار"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TCOD_ANBAR.cls"
