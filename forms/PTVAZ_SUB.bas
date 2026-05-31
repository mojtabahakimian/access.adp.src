Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11610
    RowHeight =375
    ItemSuffix =10
    Left =8040
    Top =1140
    Right =15000
    Bottom =9360
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x32b28cd4ec2ce340
    End
    RecordSource ="PTVAZ"
    Caption ="PTVAZ"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =6720
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =45
                    ColumnWidth =1417
                    Name ="PIDT"
                    ControlSource ="PIDT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =45
                            Width =1065
                            Height =240
                            Name ="Label1"
                            Caption ="شماره  پذيرش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =725
                    Width =2490
                    ColumnWidth =1080
                    TabIndex =2
                    Name ="PDATE"
                    ControlSource ="PDATE"
                    DefaultValue ="=FARSIDATE(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =725
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =1065
                    ColumnWidth =840
                    TabIndex =3
                    Name ="PTIME"
                    ControlSource ="PTIME"
                    DefaultValue ="Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1065
                            Width =540
                            Height =240
                            Name ="Label7"
                            Caption ="ساعت"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =1406
                    ColumnWidth =1417
                    TabIndex =4
                    Name ="PIDC"
                    ControlSource ="PIDC"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1410
                            Width =450
                            Height =240
                            Name ="Label9"
                            Caption ="ترتيب"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =164
                    Top =385
                    Width =4656
                    ColumnWidth =4665
                    TabIndex =1
                    Name ="PVAZ"
                    ControlSource ="PVAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"در انتظار تعمير\";2;\"در حال تعمير\";3;\"منتظر قطعه\";4;\"آماده تحويل\";5;\""
                        "غير قابل تعمير\";6;\"ارسال جهت تعمير خارج از شركت\";7;\"تحويل شد\";8;\"انصراف از"
                        " تعمير\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =385
                            Width =1185
                            Height =240
                            Name ="Label3"
                            Caption ="وضعيت دستگاه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PTVAZ_SUB.cls"
