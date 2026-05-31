Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    RowHeight =255
    ItemSuffix =11
    Left =630
    Top =2145
    Right =9375
    Bottom =5490
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xaf3dda3abe4ee240
    End
    RecordSource ="CHECK_BOOK"
    Caption ="CHECK_BOOK"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
        Begin PageHeader
            DisplayWhen =1
            Height =1134
            BackColor =-2147483633
            Name ="PageHeader"
        End
        Begin Section
            Height =2012
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =453
                    ColumnWidth =1830
                    TabIndex =1
                    Name ="DATEDB"
                    ControlSource ="DATEDB"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3620
                            Top =453
                            Width =885
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ دريافت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =793
                    ColumnWidth =1635
                    TabIndex =2
                    Name ="SER_FRM"
                    ControlSource ="SER_FRM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3845
                            Top =793
                            Width =660
                            Height =240
                            Name ="Label5"
                            Caption ="از سريال"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =1133
                    ColumnWidth =1665
                    TabIndex =3
                    Name ="SER_TO"
                    ControlSource ="SER_TO"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3860
                            Top =1133
                            Width =645
                            Height =240
                            Name ="Label7"
                            Caption ="تا سريال"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =113
                    Width =2490
                    Height =255
                    ColumnWidth =1050
                    Name ="COD_HES"
                    ControlSource ="COD_HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3470
                            Top =113
                            Width =1035
                            Height =240
                            Name ="Label1"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =510
                    Top =1757
                    Width =516
                    Height =255
                    ColumnWidth =1740
                    TabIndex =4
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"10 برگي\";2;\"25 برگي\";3;\"50 برگي\";4;\"100 برگي\""
                    ColumnWidths ="0"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3552
                            Top =1757
                            Width =360
                            Height =240
                            Name ="Label8"
                            Caption ="نوع"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =1134
            BackColor =-2147483633
            Name ="PageFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =2490
                    Height =255
                    Name ="CCODH"
                    ControlSource ="=Count([COD_HES])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3326
                            Width =1035
                            Height =240
                            Name ="Label10"
                            Caption ="شماره حساب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHECK_BOOK.cls"
