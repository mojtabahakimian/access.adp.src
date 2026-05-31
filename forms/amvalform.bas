Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14492
    RowHeight =315
    ItemSuffix =24
    Left =420
    Right =14925
    Bottom =8265
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OrderBy ="amval.barchasb"
    RecSrcDt = Begin
        0xa60a6d43910ce540
    End
    RecordSource ="amval"
    Caption ="اموال و دارائيها"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
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
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =19502
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10588
                    Top =113
                    Width =1440
                    Height =375
                    ColumnWidth =690
                    ColumnOrder =0
                    BackColor =255
                    Name ="barchasb"
                    ControlSource ="barchasb"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =113
                            Width =1815
                            Height =360
                            Name ="Label1"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4485
                    Top =570
                    Width =7545
                    Height =375
                    ColumnWidth =3690
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =0
                    ForeColor =16777215
                    Name ="aname"
                    ControlSource ="aname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =570
                            Width =525
                            Height =360
                            Name ="Label3"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10335
                    Top =1485
                    Height =375
                    ColumnWidth =1875
                    ColumnOrder =4
                    TabIndex =3
                    Name ="rate"
                    ControlSource ="rate"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12135
                            Top =1485
                            Width =1905
                            Height =360
                            Name ="Label7"
                            Caption ="ضريب محاسبه استهلاك "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10830
                    Top =2385
                    Width =1200
                    Height =375
                    ColumnOrder =5
                    TabIndex =6
                    Name ="khdate"
                    ControlSource ="khdate"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =2386
                            Width =1020
                            Height =360
                            Name ="Label11"
                            Caption ="تاريخ خريد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =2835
                    Width =2490
                    Height =375
                    ColumnOrder =6
                    TabIndex =7
                    Name ="price"
                    ControlSource ="price"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =2841
                            Width =1050
                            Height =360
                            Name ="Label13"
                            Caption ="قيمت خريد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =3735
                    Width =2490
                    Height =375
                    TabIndex =9
                    Name ="arzesh"
                    ControlSource ="arzesh"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =3736
                            Width =1185
                            Height =360
                            Name ="Label17"
                            Caption ="ارزش دفتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =285
                    Top =4635
                    Width =11745
                    Height =375
                    TabIndex =11
                    Name ="Description"
                    ControlSource ="Description"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =4635
                            Width =630
                            Height =360
                            Name ="Label21"
                            Caption ="شرح"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10329
                    Top =1020
                    Height =375
                    ColumnWidth =1755
                    ColumnOrder =3
                    TabIndex =2
                    Name ="calmethod"
                    ControlSource ="calmethod"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مستقيم\";2;\"نزولي\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =1022
                            Width =1785
                            Height =360
                            Name ="Label5"
                            Caption ="روش محاسبه استهلاك"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10329
                    Top =4185
                    Height =375
                    TabIndex =10
                    BackColor =6723891
                    ForeColor =10092543
                    Name ="astatus"
                    ControlSource ="astatus"
                    RowSourceType ="Value List"
                    RowSource ="1;\"درحال استفاده\";2;\"راكد\";3;\"مفقودي\";4;\"در دست تعمير\";5;\"فروخته شده\";"
                        "10;\"نامشخص\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =4185
                            Width =825
                            Height =360
                            Name ="Label19"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =5875
                    Top =1962
                    Width =6171
                    Height =315
                    ColumnWidth =3915
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="hesab"
                    ControlSource ="hesab"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;4536;567"
                    Tag ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12142
                            Top =1961
                            Width =2235
                            Height =240
                            Name ="Label22"
                            Caption =" حساب معادل در حسابداري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3750
                    Top =1950
                    Width =2091
                    Height =315
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="hesab2"
                    ControlSource ="hesab"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3750
                    Top =3289
                    Width =8280
                    Height =405
                    ColumnWidth =2085
                    ColumnOrder =2
                    TabIndex =8
                    Name ="place"
                    ControlSource ="place"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT place FROM amval GROUP BY place ORDER BY place"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12132
                            Top =3285
                            Width =1245
                            Height =360
                            Name ="Label15"
                            Caption ="موقعيت مكاني"
                        End
                    End
                End
                Begin Image
                    SpecialEffect =2
                    SizeMode =3
                    PictureType =1
                    Left =56
                    Top =56
                    Width =3581
                    Height =4529
                    Name ="PIC"
                    ObjectPalette = Begin
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
                    End
                    Picture ="F:\\software\\prg\\SAFIR\\Files\\11.bmp"
                End
                Begin Image
                    SpecialEffect =2
                    SizeMode =1
                    PictureType =1
                    Left =113
                    Top =5102
                    Width =14231
                    Height =14175
                    Name ="picl"
                    ObjectPalette = Begin
                        0x000300010000000020000000400000006000000080000000a0000000c0000000 ,
                        0xe00000000020000020200000402000006020000080200000a0200000c0200000 ,
                        0xe02000000040000020400000404000006040000080400000a0400000c0400000 ,
                        0xe04000000060000020600000406000006060000080600000a0600000c0600000 ,
                        0xe06000000080000020800000408000006080000080800000a0800000c0800000 ,
                        0xe080000000a0000020a0000040a0000060a0000080a00000a0a00000c0a00000 ,
                        0xe0a0000000c0000020c0000040c0000060c0000080c00000a0c00000c0c00000 ,
                        0xe0c0000000e0000020e0000040e0000060e0000080e00000a0e00000c0e00000 ,
                        0xe0e000000000400020004000400040006000400080004000a0004000c0004000 ,
                        0xe00040000020400020204000402040006020400080204000a0204000c0204000 ,
                        0xe02040000040400020404000404040006040400080404000a0404000c0404000 ,
                        0xe04040000060400020604000406040006060400080604000a0604000c0604000 ,
                        0xe06040000080400020804000408040006080400080804000a0804000c0804000 ,
                        0xe080400000a0400020a0400040a0400060a0400080a04000a0a04000c0a04000 ,
                        0xe0a0400000c0400020c0400040c0400060c0400080c04000a0c04000c0c04000 ,
                        0xe0c0400000e0400020e0400040e0400060e0400080e04000a0e04000c0e04000 ,
                        0xe0e040000000800020008000400080006000800080008000a0008000c0008000 ,
                        0xe00080000020800020208000402080006020800080208000a0208000c0208000 ,
                        0xe02080000040800020408000404080006040800080408000a0408000c0408000 ,
                        0xe04080000060800020608000406080006060800080608000a0608000c0608000 ,
                        0xe06080000080800020808000408080006080800080808000a0808000c0808000 ,
                        0xe080800000a0800020a0800040a0800060a0800080a08000a0a08000c0a08000 ,
                        0xe0a0800000c0800020c0800040c0800060c0800080c08000a0c08000c0c08000 ,
                        0xe0c0800000e0800020e0800040e0800060e0800080e08000a0e08000c0e08000 ,
                        0xe0e080000000c0002000c0004000c0006000c0008000c000a000c000c000c000 ,
                        0xe000c0000020c0002020c0004020c0006020c0008020c000a020c000c020c000 ,
                        0xe020c0000040c0002040c0004040c0006040c0008040c000a040c000c040c000 ,
                        0xe040c0000060c0002060c0004060c0006060c0008060c000a060c000c060c000 ,
                        0xe060c0000080c0002080c0004080c0006080c0008080c000a080c000c080c000 ,
                        0xe080c00000a0c00020a0c00040a0c00060a0c00080a0c000a0a0c000c0a0c000 ,
                        0xe0a0c00000c0c00020c0c00040c0c00060c0c00080c0c000a0c0c000c0c0c000 ,
                        0xe0c0c00000e0c00020e0c00040e0c00060e0c00080e0c000a0e0c000c0e0c000 ,
                        0xe0e0c00000000000
                    End
                    Picture ="F:\\software\\prg\\SAFIR\\Files\\11.bmp"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3810
                    Top =510
                    Width =615
                    Height =465
                    TabIndex =12
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =3683
                    Top =59
                    Width =2106
                    Height =315
                    TabIndex =13
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5846
                            Top =56
                            Width =525
                            Height =300
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="كاربر:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "amvalform.cls"
