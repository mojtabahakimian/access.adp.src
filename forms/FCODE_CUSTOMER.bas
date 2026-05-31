Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11279
    RowHeight =345
    ItemSuffix =50
    Left =8640
    Top =1320
    Right =19920
    Bottom =10035
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x2fce5407f101e640
    End
    RecordSource ="SELECT N_KOL, CUST_COD, NUMBER, TNUMBER, NAME, TOZIH, BED_BES, ADDRESS, TEL, COD"
        "E_E, ECODE, PCODE, IYALAT, CITY, MCODEM, MOBILE, ROUTE_NAME, RTRIM(CAST(N_KOL AS"
        " nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER A"
        "S nvarchar)) AS HES, Longitude, Latitude, OSTANID, SHAHRID, tob FROM TDETA_HES W"
        "HERE (N_KOL = 124) AND (NUMBER = 1)"
    Caption ="تعريف  مشتري جديد"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =8730
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8217
                    Top =120
                    Width =685
                    Height =300
                    ColumnOrder =0
                    TabIndex =1
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    DefaultValue ="0"
                    Tag ="1374"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9054
                            Top =120
                            Width =225
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
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4873
                    Top =460
                    Width =4020
                    Height =300
                    ColumnWidth =3915
                    ColumnOrder =1
                    Name ="NAMES"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9054
                            Top =457
                            Width =300
                            Height =285
                            Name ="Label3"
                            Caption ="نام "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2263
                    Top =3607
                    Width =6636
                    Height =1410
                    ColumnWidth =4035
                    ColumnOrder =4
                    TabIndex =10
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9052
                            Top =3600
                            Width =660
                            Height =285
                            Name ="Label7"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2202
                    Top =809
                    Width =6696
                    Height =300
                    ColumnWidth =5460
                    ColumnOrder =2
                    TabIndex =2
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9053
                            Top =795
                            Width =450
                            Height =285
                            Name ="Label9"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4857
                    Top =1168
                    Width =4041
                    Height =300
                    ColumnOrder =3
                    TabIndex =3
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9053
                            Top =1155
                            Width =360
                            Height =285
                            Name ="Label10"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin UnboundObjectFrame
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =737
                    Top =283
                    Width =1065
                    Height =1590
                    TabIndex =11
                    Name ="OLEUnbound12"
                    OleData = Begin
                        0x00660000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000010000000000000000100000 ,
                        0x0200000001000000feffffff0000000000000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffdffffff31000000fefffffffeffffff050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x28000000290000002a0000002b0000002c0000002d0000002e0000002f000000 ,
                        0x30000000fefffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff020000000a00030000000000c0000000 ,
                        0x0000004600000000000000000000000040ed4a322ac1db010300000040010000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000014000000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010100000004000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000010000004d000000 ,
                        0x0000000001004f006c006500310030004e006100740069007600650000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000004000000c4590000 ,
                        0x00000000feffffff02000000fefffffffefffffffeffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000100feff030a0000ffffffff0a00030000000000c000000000000046 ,
                        0x0700000050427275736800070000005042727573680007000000504272757368 ,
                        0x00f439b271000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000001000000480000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000038000000000000000100000060006000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000c0590000424da6590000000000003600000028000000470000006a00 ,
                        0x0000010018000000000070590000c40e0000c40e00000000000000000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebe000000ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00000000ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00e0c000000000e0c000ffff000000 ,
                        0x00bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00000000e0c000ffff00ffff00000000bebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xe0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00000000e0c000ffff00ffff00000000bebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000ffff00ffff00ffff00ffff00000000e0c000ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xe0c000ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000ffff ,
                        0x00ffff00ffff00ffff00000000e0c000ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00000000e0c000e0c000ffff ,
                        0x00ffff00000000bebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe000000ffff00ffff00ff ,
                        0xff00000000e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00000000e0c000ffff00ffff00000000be ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebe000000ffff00ffff00ffff00000000ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00000000e0c000ffff00ffff00000000bebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe008000bebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebe000000ffff00ffff00ffff00000000ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00e0 ,
                        0xc000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0x000000e0c000ffff00ffff00ffff00000000bebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebe008000008000bebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe0000 ,
                        0x00ffff00e0c000000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00e0c000ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000e0c000e0c0 ,
                        0x00ffff00ffff00000000bebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebe00800000ff00008000008000bebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe008000e0c000000000000000ff ,
                        0xe0c0000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00e0c000ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00000000e0c000e0c000ffff0000 ,
                        0x0000bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbe00800000ff0000ff00008000008000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebe008000008000000000ffe0c0ffe0c0ffe0c0ffe0c0000000 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00e0c000e0c000000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00000000e0c000e0c000ffff00000000bebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebe00800000ff0000 ,
                        0xff0000ff0000ff00008000008000008000bebebebebebebebebebebebe008000 ,
                        0x00ff00000000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0000000ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00e0c000000000e0 ,
                        0xc000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00000000e0c000e0c000000000bebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebe00800000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff00008000bebebebebebebebebe00800000ff0000ff00000000ffe0 ,
                        0xc0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0000000ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00000000e0c000e0c000ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00000000e0c000e0c000000000bebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebe00800000ff0000ff0000ff0000ff0000ff0000ff000080 ,
                        0x00bebebe00800000800000ff0000ff00000000ffe0c0ffe0c0ffe0c0ffe0c0ff ,
                        0xe0c0ffe0c0ffe0c0ffe0c0000000ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00e0c000000000e0c000e0c000ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000e0 ,
                        0xc000e0c000bebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbe00800000ff0000ff0000ff0000ff0000ff0000ff0000800000800000ff0000 ,
                        0xff0000ff0000ff00000000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0 ,
                        0xffe0c0000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff000000 ,
                        0x00000000e0c000e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000e0c000bebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebe00800000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000ff0000ff00 ,
                        0x000000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c00000 ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000000000e0c000e0 ,
                        0xc000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00000000e0c000bebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebe00800000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000800000ff0000ff0000ff0000ff0000ff00000000000000ffe0 ,
                        0xc0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0000000ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00e0c000000000ffff00000000e0c000ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebe00800000ff0000ff0000ff0000ff0000ff0000ff000080 ,
                        0x0000ff0000ff0000ff0000ff0000ff0000000080a0e0000000ffe0c0ffe0c0ff ,
                        0xe0c0ffe0c0ffe0c0ffe0c0ffe0c0000000ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00000000ffff00ffff00000000e0c000ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00000000e0c000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebe0080 ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000800000800000ff0000ff0000 ,
                        0xff0000ff0000ff0000000080a0e080a0e0000000ffe0c0ffe0c0ffe0c0ffe0c0 ,
                        0xffe0c0ffe0c0ffe0c0000000ffff00ffff00ffff00ffff00e0c000000000ffff ,
                        0x00ffff00000000e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000e0c000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe00800000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000800000800000ff0000ff0000ff0000ff00000000 ,
                        0x80a0e080a0e080a0e080a0e0000000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0 ,
                        0xc0000000ffff00ffff00ffff00e0c000000000ffff00ffff00ffff00000000e0 ,
                        0xc000e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebe00800000ff0000ff0000ff0000ff0000ff00 ,
                        0x00800000ff0000800000ff0000ff0000ff0000ff0000000080a0e080a0e080a0 ,
                        0xe080a0e0000000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0000000ffff00ff ,
                        0xff00ffff00000000ffff00ffff00ffff00ffff00e0c000000000e0c000ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebe00800000ff0000ff0000ff0000ff0000ff0000800000ff000080 ,
                        0x0000ff0000ff0000ff0000000080a0e080a0e080a0e080a0e080a0e080a0e000 ,
                        0x0000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c0000000ffff00ffff00e0c000000000 ,
                        0xffff00ffff00ffff00ffff00ffff00000000e0c000e0c000ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00e0c000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebe00800000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000800000ff0000800000ff0000ff0000 ,
                        0x000080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000ffe0c0ffe0c0 ,
                        0xffe0c0ffe0c0ffe0c0000000ffff00e0c000000000ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00e0c000000000e0c000ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebe00800000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000800000ff0000800000ff0000ff0000000080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e0000000ffe0c0ffe0c0ffe0c0ffe0 ,
                        0xc0ffe0c0000000000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00000000e0c000e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebe00800000ff0000ff0000ff0000ff0000ff0000ff00008000 ,
                        0x00ff0000ff0000800000000000000080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e0000000ffe0c0ffe0c0ffe0c0ffe0c0ffe0c000000000 ,
                        0x0000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xe0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0x00800000ff0000ff0000ff0000ff0000ff0000ff0000800000ff0000ff000000 ,
                        0x0080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e0000000ffe0c0ffe0c0ffe0c0ffe0c0000000000000000000000000ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000e0c000ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00000000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebe00800000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000800000ff0000ff0000000080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000ffe0c0 ,
                        0xffe0c0ffe0c0000000a000ff000000ffff00000000ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00000000ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebe00800000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000800000ff0000ff0000000080a0e080a0e000000000000080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e0000000ffe0c0ffe0c0ffe0c00000 ,
                        0x00a000ffa000ff000000000000000000ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00000000e0c000ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebe00800000ff0000ff0000ff0000ff0000ff0000ff0000ff00008000008000 ,
                        0x00ff0000ff0000000080a0e000000080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e0000000ffe0c0000000a000ffa000ffa000ff00 ,
                        0x0000000000000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00000000e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebe00800000800000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000800000800000800000ff000080 ,
                        0x0000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e0000000000000a000ffa000ffa000ffa000ff000000000000ffe0e0 ,
                        0x000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00e0c0 ,
                        0x00000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00000000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe00000000800000800000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000800000800000ff0000ff0000800000000080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000 ,
                        0x000000a000ffa000ffa000ffa000ff000000ffff00000000000000ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00e0c000000000ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000ffff ,
                        0xff00800000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0x800000800000ff0000ff0000000080a0e080a0e080a0e080a0e080a0e0000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000ffff00000000a000 ,
                        0xffa000ffa000ff000000ffff00000000ffe0e0000000ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00000000ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000ffffff00800000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000800000ff00 ,
                        0x00ff0000000080a0e080a0e080a0e000000000000080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e0000000ffff00000000a000ffa000ffa000ff00 ,
                        0x0000ffff00000000ffe0e0000000ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00000000ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000ffffffffffff00000000000000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000800000ff0000ff0000000080a0 ,
                        0xe000000000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e0000000008000000000a000ffa000ff000000ffff00000000 ,
                        0xffe0e0ffe0e0000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00e0c000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00000000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000ffffff00000080a0e080a0e000000000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000800000ff0000ff0000ff0000000000000080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e0000000000000a000ffa000ffa000ff000000000000ffe0e0ffe0e00000 ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000ffff ,
                        0xff00000080a0e080a0e080a0e000000000ff0000ff0000ff0000ff0000ff0000 ,
                        0x800000ff0000ff0000ff0000000000000080a0e080a0e080a0e080a0e080a0e0 ,
                        0x00000000000000000080a0e080a0e080a0e080a0e080a0e080a0e00000000000 ,
                        0x00a000ffa000ffa000ff000000000000ffe0e0ffe0e0ffe0e0000000ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000ffffff00000000000080 ,
                        0xa0e080a0e000000000ff0000ff0000ff0000ff0000ff0000800000ff0000ff00 ,
                        0x00ff0000ff0000000080a0e080a0e080a0e000000000000080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e0000000ffff00000000a000ffa0 ,
                        0x00ff000000000000ffe0e0ffe0e0ffe0e0ffe0e0000000ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000ffffff00000080a0e000000080a0e0000000 ,
                        0x00ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000ff0000000080a0 ,
                        0xe080a0e080a0e000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e0000000ffff00000000a000ffa000ff000000000000 ,
                        0xffe0e0ffe0e0ffe0e0ffe0e0ffe0e0000000ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00000000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000ffffff00000080a0e000000000000000000000ff0000ff0000ff ,
                        0x0000ff0000800000ff0000ff0000ff0000ff0000000080a0e080a0e000000080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e000000000000080a0e0 ,
                        0x80a0e0000000000000ffff00000000a000ffa000ff000000ffe0e0ffe0e0ffe0 ,
                        0xe0ffe0e0ffe0e0ffe0e0000000000000ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000ffff ,
                        0xffffffff00000080a0e080a0e080a0e000000000ff0000ff0000800000ff0000 ,
                        0xff0000ff0000ff0000ff0000000080a0e000000080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e000000000000000000000000080a0e080a0e080a0e0000000ffe0 ,
                        0xe0000000000000a000ffa000ffa000ff000000ffe0e0ffe0e0ffe0e0ffe0e0ff ,
                        0xe0e0ffe0e0ffe0e0000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000ffffffffffff00000080 ,
                        0xa0e000000000000000000000000000800000800000ff0000ff0000ff0000ff00 ,
                        0x00800000000000000080a0e080a0e080a0e080a0e080a0e000000000000000ff ,
                        0x0000ff0000ff0000000000000080a0e080a0e0000000ffe0e0000000000000a0 ,
                        0x00ffa000ffa000ff000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0 ,
                        0xffe0e0000000000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00ffff00000000bebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe00000000800000800000ff0000000080a0e080a0e0 ,
                        0x80a0e080a0e000000000000000ff0000ff0000ff0000800000ff0000000080a0 ,
                        0xe080a0e080a0e080a0e080a0e000000000800000ff0000ff0000ff0000ff0000 ,
                        0xff0000000080a0e0000000000000ffe0e0ffe0e0000000a000ffa000ffa000ff ,
                        0xa000ff000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0 ,
                        0xe0000000ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00ff ,
                        0xff00000000bebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebe00000000000080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e000000000ff0000800000ff0000ff0000000080a0e080a0e080a0e080 ,
                        0xa0e000000000ff0000ff0000800000ff0000ff0000ff00008000008000000000 ,
                        0x000000000000ffe0e0ffe0e0000000a000ffa000ffa000ffa000ff000000ffe0 ,
                        0xe0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0000000ff ,
                        0xff00ffff00ffff00ffff00ffff00ffff00ffff00ffff00000000bebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbe00000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e000 ,
                        0x000000ff0000ff0000ff0000000080a0e080a0e080a0e000000000ff0000ff00 ,
                        0x00ff0000ff0000800000ff0000ff0000800000ff00000000ffff00000000ffe0 ,
                        0xe0ffe0e0000000a000ffa000ffa000ffa000ff000000ffe0e0ffe0e0ffe0e0ff ,
                        0xe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0000000ffff00ffff00ffff00 ,
                        0xffff00ffff00ffff00ffff00ffff00000000bebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebe00000080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e000000000ff00 ,
                        0x00ff0000000080a0e080a0e000000000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000800000ff0000800000ff0000ff00008000008000000000ffe0e0000000a0 ,
                        0x00ffa000ffa000ffa000ff000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0 ,
                        0xffe0e0ffe0e0ffe0e0ffe0e0ffe0e0000000ffff00ffff00ffff00ffff00ffff ,
                        0x00ffff00ffff00000000bebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebe00000080a0e0000000000000 ,
                        0x00000000000000000000000000000080a0e000000000ff0000ff0000000080a0 ,
                        0xe000000000800000ff0000ff0000ff0000ff0000ff0000ff0000ff0000800000 ,
                        0x800000ff0000ff00008000ffff00000000ffe0e0000000a000ffa000ffa000ff ,
                        0xa000ff000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0 ,
                        0xe0ffe0e0ffe0e0ffe0e0000000ffff00ffff00ffff00ffff00ffff00000000be ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebe00000000000000000080a0e080a0e080a0e00000000000 ,
                        0x0000ff0000800000000000ff0000ff0000ff0000ff0000000000ff0000ff0000 ,
                        0x800000ff0000ff0000ff0000ff0000ff0000ff0000ff0000800000ff0000ff00 ,
                        0x008000ffff00000000ffe0e0ffe0e0000000a000ffa000ffa000ff000000ffe0 ,
                        0xe0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ff ,
                        0xe0e0000000ffff00ffff00ffff00ffff00ffff00000000bebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbe00000080a0e080a0e080a0e080a0e080a0e080a0e080a0e000000000000000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000800000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000ff00008000000000ffff000000 ,
                        0x00ffe0e0ffe0e0ffe0e0000000000000000000ffe0e0ffe0e0ffe0e0ffe0e0ff ,
                        0xe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0000000 ,
                        0xffff00ffff00ffff00000000bebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebe00000080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e000000000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff00008000bebebe000000ffff00000000ffe0e0ff ,
                        0xe0e0ffe0e0ffe0e0000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0 ,
                        0xffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0000000ffff00ffff000000 ,
                        0x00bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebe00000080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e000000000ff0000ff0000ff0000ff ,
                        0x0000ff0000ff0000ff0000800000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff00008000bebebebebebe000000ffff00000000ffe0e0ffe0e0ffe0e0ffe0e0 ,
                        0xffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e00000 ,
                        0x00000000ffe0e0ffe0e0ffe0e0ffe0e0000000000000000000bebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebe00000080a0e00000000000000000000000000000000000 ,
                        0x0080a0e080a0e080a0e000000000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff0000800000800000ff0000ff0000ff0000ff0000ff0000ff00008000bebebe ,
                        0xbebebebebebe000000000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e0ffe0 ,
                        0xe0ffe0e0ffe0e0ffe0e0ffe0e0ffe0e000000000000080a0e080a0e000000000 ,
                        0x0000000000000000ffffff000000000000000000bebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebe00000080a0e080a0e080a0e080a0e000000000000000000080a0e080 ,
                        0xa0e000000000ff0000ff0000ff0000ff0000ff0000ff0000800000ff00008000 ,
                        0x00ff0000ff0000ff0000ff0000ff0000ff00008000bebebebebebebebebebebe ,
                        0xbe000000000000ffe0e0ffe0e0ffe0e0ffe0e0ffe0e000000000000000000000 ,
                        0x000000000000000080a0e080a0e080a0e080a0e080a0e080a0e0000000000000 ,
                        0x00000080a0e080a0e080a0e0000000000000bebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebe00000080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e000000000000000000000000000ff00 ,
                        0x00ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000800000ff0000ff ,
                        0x0000ff0000ff0000ff00008000bebebebebebebebebebebebe000000000000ff ,
                        0xe0e0ffe0e0ffe0e0ffe0e0ffe0e000000080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e000000000000080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e0000000000000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebe00000080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e000000000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000800000ff0000ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000 ,
                        0x8000008000bebebebebebebebebebebebebebebe000000000000ffe0e0ffe0e0 ,
                        0xffe0e0ffe0e000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e0000000bebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebe00000080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e000000000ff0000ff0000ff0000ff0000ff0000800000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000800000ff0000ff0000800000ff00008000 ,
                        0xbebebebebebebebebebebebebebebe000000ffe0e0ffe0e0ffe0e0ffe0e00000 ,
                        0x0080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x000000bebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebe00000080a0e080a0e080a0e000000080a0e080a0e080a0e080a0e000 ,
                        0x000000ff0000ff0000ff0000ff0000800000800000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000800000800000ff0000800000ff00008000bebebebebebebebe ,
                        0xbebebebebebebebebebe000000ffe0e0ffe0e000000080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebe00000000 ,
                        0x000000000000000000800000000000000080a0e080a0e000000000ff0000ff00 ,
                        0x00ff0000ff0000800000800000ff0000ff0000ff0000ff0000ff0000ff000080 ,
                        0x0000ff0000800000800000ff00008000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebe000000ffe0e000000080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebe00800000ff0000ff00 ,
                        0x00ff0000800000800000000000000000000000ff0000ff0000ff0000800000ff ,
                        0x0000800000ff0000ff0000ff0000ff0000ff0000ff0000800000ff0000ff0000 ,
                        0x800000ff00008000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0x00000000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e0000000bebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebe00800000ff0000ff0000ff0000ff0000ff000080 ,
                        0x0000ff0000800000800000ff0000ff0000ff0000800000ff0000800000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000ff00008000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebe00000000000080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e0000000bebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebe00800000ff0000ff0000ff0000ff0000ff0000800000ff0000ff0000 ,
                        0xff0000800000ff0000ff0000800000ff0000800000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000800000ff0000ff0000ff0000ff00008000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe00000080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe0000000bebebebebebebebebebebebebebebe000000bebebebebebe00800000 ,
                        0x800000ff0000ff0000ff0000ff0000800000ff0000ff0000ff0000800000ff00 ,
                        0x00ff0000800000ff0000800000ff0000ff0000ff0000ff0000ff000080000080 ,
                        0x0000ff0000ff0000ff0000ff00008000bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe00000080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebe008000008000 ,
                        0x00800000ff0000800000ff0000ff0000ff0000ff0000800000800000ff0000ff ,
                        0x0000800000ff0000ff0000ff0000800000800000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff00008000bebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebe00000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e000000080a0e000000000000080a0e080a0e0000000bebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebe0080000080 ,
                        0x0000ff0000ff0000ff0000ff0000800000800000ff0000ff0000800000ff0000 ,
                        0xff0000800000800000800000ff0000ff0000ff0000ff0000ff0000ff00008000 ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe0000 ,
                        0x0080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000000000 ,
                        0x4080c000000080a0e080a0e0000000bebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe00800000800000ff0000 ,
                        0xff0000ff0000800000800000ff0000800000ff0000800000800000ff0000ff00 ,
                        0x00800000ff0000ff0000ff0000ff0000ff00008000008000bebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebe00000080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e00000000000000000000000004080c000000000000080a0 ,
                        0xe080a0e0000000bebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebe00800000ff0000800000ff0000ff0000ff00008000 ,
                        0x00ff0000ff0000800000800000ff0000ff0000ff0000ff0000800000ff0000ff ,
                        0x0000ff00008000008000bebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebe00000080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e000000080a0 ,
                        0xe00000004080c00000000000004080c000000000000000000080a0e0000000be ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebe008000 ,
                        0x00800000ff0000ff0000ff0000800000ff0000ff0000800000ff0000ff000080 ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000800000ff00008000bebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebe00000080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e00000000000000000004080c040 ,
                        0x80c00000004080c0000000000000000000000000000000bebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebe00800000ff0000ff0000ff0000ff ,
                        0x0000ff0000800000ff0000ff0000ff0000800000ff0000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff00008000bebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe0000 ,
                        0x0080a0e000000000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e00000004080c00000000000004080c04080c00000004080c0 ,
                        0x0000000000004080c000000080a0e0000000bebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebe00800000ff0000ff0000ff0000ff0000800000 ,
                        0xff0000ff0000ff0000800000ff0000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff00008000bebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe00000080a0e080a0e000 ,
                        0x000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x0000004080c04080c04080c00000004080c00000004080c00000000000004080 ,
                        0xc000000080a0e0000000bebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebe00800000ff0000ff0000ff0000ff0000ff0000800000ff0000ff00 ,
                        0x00800000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff00008000008000bebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe00000080a0e080a0e080a0e000000080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e00000004080c04080 ,
                        0xc04080c04080c04080c00000000000000000000000004080c000000080a0e000 ,
                        0x0000bebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0x00800000ff0000ff0000ff0000ff0000800000ff0000ff0000800000ff0000ff ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000 ,
                        0xff00008000bebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0x00000080a0e080a0e080a0e080a0e000000080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e000000000000000000000000000000000000040 ,
                        0x80c04080c00000000000004080c04080c0000000000000000000bebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebe00800000ff0000ff ,
                        0x0000ff0000ff0000800000ff0000ff0000ff0000800000ff0000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00008000bebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebe00000080a0e080a0e080a0 ,
                        0xe080a0e080a0e000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e000000080a0e080a0e080a0e080a0e080a0e0000000000000000000 ,
                        0x00000000000000000000000080a0e080a0e0000000bebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebe00800000ff0000ff0000ff0000 ,
                        0xff0000800000ff0000ff0000800000ff0000ff0000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff0000ff0000ff00008000bebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebe00000080a0e080a0e080a0e080a0e080a0e000 ,
                        0x000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e000000000000000000080a0 ,
                        0xe080a0e080a0e080a0e0000000bebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebe00800000ff0000ff0000ff0000ff0000800000ff00 ,
                        0x00ff0000800000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff ,
                        0x0000ff00008000bebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebe00000080a0e080a0e080a0e080a0e080a0e000000080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e000000080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e000000080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e0000000bebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebe00800000ff0000ff0000ff0000ff0000800000ff0000ff0000ff000080 ,
                        0x0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff0000ff00008000be ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebe000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e000000080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e000000080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebe00800000ff ,
                        0x0000ff0000ff0000ff0000ff0000800000ff0000ff0000800000ff0000ff0000 ,
                        0xff0000ff0000ff0000ff0000ff0000ff00008000bebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebe00000080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e000000080a0e080a0e080a0e080a0e080a0e080a0e000000080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e0000000bebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe00800000ff0000ff0000 ,
                        0xff0000ff0000800000800000800000ff0000800000ff0000ff0000ff0000ff00 ,
                        0x00ff0000ff0000ff00008000bebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebe00000080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000 ,
                        0x80a0e080a0e080a0e000000000000000000080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e0000000bebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebe00800000ff0000ff0000ff00008000bebebe ,
                        0xbebebebebebe00800000ff0000800000ff0000ff0000ff0000ff0000ff000080 ,
                        0x00bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebe00000080a0e080a0e080a0e080a0e080a0e080a0e0000000000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e000000080a0e080a0e00000 ,
                        0x00f0f0f0f0f0f0f0f0f000000080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e0000000bebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebe00800000ff0000ff00008000bebebebebebebebebebebebebebe ,
                        0xbe00800000ff0000800000ff0000ff0000ff0000ff00008000bebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebe000000 ,
                        0x00000080a0e080a0e080a0e080a0e00000004080c04080c000000080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e000000080a0e0000000f0f0f0f0f0f0f0f0f0f0 ,
                        0xf0f0f0f0f000000080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebe0080 ,
                        0x00008000008000bebebebebebebebebebebebebebebebebebebebebe00800000 ,
                        0x800000ff0000ff0000ff0000ff00008000bebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebe00000000000000000080a0e080a0 ,
                        0xe080a0e080a0e00000004080c04080c000000080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e0000000000000000000000000f0f0f0f0f0f0f0f0f0000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e0000000000000bebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebe008000bebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebe00800000800000ff00 ,
                        0x00ff0000ff00008000bebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebe00000000000000000000000000000080a0e080a0e04080c040 ,
                        0x80c04080c000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e0000000000000000000f0f0f0f0f0f0f0f0f000000080a0e080a0e080a0 ,
                        0xe080a0e080a0e0000000000000bebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebe00800000ff00008000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe00 ,
                        0x00000000000000000000000000000000004080c00000004080c04080c0000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e00000000000 ,
                        0x00000000f0f0f0f0f0f0f0f0f000000080a0e080a0e080a0e080a0e0000000f0 ,
                        0xf0f0000000bebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebe008000008000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebe000000000000000000 ,
                        0x4080c00000004080c04080c00000004080c04080c000000080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e0000000000000000000f0f0f0f0 ,
                        0xf0f0f0f0f000000080a0e080a0e080a0e080a0e0000000f0f0f0000000bebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe008000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebe0000000000004080c00000000000004080 ,
                        0xc04080c00000000000004080c000000080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e0000000f0f0f0f0f0f0f0f0f0f0f0f0000000 ,
                        0x80a0e080a0e080a0e0000000f0f0f0000000000000bebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebe0000000000004080c04080c00000004080c04080c000000000 ,
                        0x00004080c000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e0000000000000f0f0f0f0f0f0f0f0f0f0f0f0f0f0f000000080a0e080a0 ,
                        0xe0000000f0f0f0000000000000bebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe00 ,
                        0x00000000004080c04080c00000004080c04080c04080c04080c04080c0000000 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e00000004080c04080 ,
                        0xc0000000000000f0f0f0f0f0f0f0f0f000000080a0e0000000f0f0f000000040 ,
                        0x80c04080c0bebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebe0000004080c04080c0 ,
                        0x4080c00000004080c04080c04080c04080c04080c000000080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e00000004080c04080c04080c04080c000 ,
                        0x0000f0f0f0000000000000000000f0f0f00000004080c04080c04080c0bebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebe0000004080c04080c04080c04080c00000 ,
                        0x004080c04080c04080c04080c000000080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e00000004080c04080c04080c04080c04080c0000000000000 ,
                        0x00000080a0e00000004080c04080c04080c04080c0bebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebe0000004080c04080c04080c04080c00000004080c04080c040 ,
                        0x80c000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e00000004080c04080c04080c04080c04080c04080c00000000000004080 ,
                        0xc04080c04080c04080c0000000bebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe00 ,
                        0x00004080c00000004080c04080c04080c00000004080c000000000000080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e00000 ,
                        0x000000004080c04080c04080c04080c04080c00000004080c04080c04080c000 ,
                        0x0000bebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebe000000000000 ,
                        0x4080c04080c04080c000000000000000000080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e000000040 ,
                        0x80c04080c04080c04080c00000004080c04080c0000000000000bebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebe0000000000004080c04080c04080 ,
                        0xc04080c000000000000080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e00000000000004080c0 ,
                        0x4080c00000004080c000000080a0e0000000bebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebe0000000000000000004080c04080c04080c000000080 ,
                        0xa0e000408080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e00000000000000000000000 ,
                        0x0080a0e0000000bebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebe0000000000004080c04080c000000000000080a0e000408080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000be ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebe000000 ,
                        0x0000000000004080c000000000408080a0e080a0e000408080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebe0000004080c00000000000 ,
                        0x0080a0e080a0e000408080a0e000408080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e0000000bebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebe00000000000000000080a0e080a0e000 ,
                        0x408080a0e000408080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe0000000bebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebe000000000000bebebe00000080a0e000408080a0e080a0e0 ,
                        0x00408080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebe00000080a0e080a0e000408080a0e080a0e000408080a0 ,
                        0xe080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e0000000bebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebe00000080a0e000408080a0e080a0e080a0e000408080a0e080a0e080 ,
                        0xa0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0 ,
                        0x80a0e0000000bebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebe00000080 ,
                        0xa0e080a0e000408080a0e080a0e080a0e000408080a0e080a0e080a0e080a0e0 ,
                        0x80a0e080a0e080a0e080a0e080a0e080a0e080a0e080a0e0000000bebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebe00000080a0e080a0e0 ,
                        0x00408080a0e080a0e080a0e000408080a0e080a0e080a0e080a0e080a0e080a0 ,
                        0xe080a0e080a0e080a0e080a0e0004080bebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000bebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebe00000080a0e080a0e000408080a0 ,
                        0xe080a0e080a0e000408000408080a0e080a0e080a0e080a0e080a0e080a0e000 ,
                        0x4080004080bebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebe000000bebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebe00000080a0e080a0e000408080a0e080a0e080 ,
                        0xa0e080a0e0004080004080004080004080004080004080004080bebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebe000000bebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebe00000000000000000000408080a0e080a0e080a0e080a0e0 ,
                        0x80a0e0004080004080004080000000bebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebe000000bebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebe0000000040800040800040800040800040800000000000 ,
                        0x00bebebebebebebebebebebebebebebebebebebebebebebebebebebebebebebe ,
                        0xbebebebebebebebebebebe000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000001004f006c006500310030004900740065006d004e0061006d006500 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e0002010300000005000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000300000005000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000400000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Class ="Paint.Picture"
                    OLEClass ="Paint"
                End
                Begin CheckBox
                    OverlapFlags =87
                    ReadingOrder =1
                    Left =8505
                    Top =3348
                    Width =395
                    Height =248
                    ColumnOrder =8
                    TabIndex =9
                    Name ="Check13"
                    ControlSource ="CODE_E"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9045
                            Top =3258
                            Width =705
                            Height =285
                            Name ="Label14"
                            Caption =" غير فعال"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3231
                    Top =56
                    Width =411
                    Height =300
                    ColumnOrder =11
                    TabIndex =12
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    DefaultValue ="1"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =1516
                    Width =2665
                    Height =300
                    ColumnOrder =5
                    TabIndex =4
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9059
                            Top =1515
                            Width =870
                            Height =240
                            Name ="Label16"
                            Caption ="كد اقتصادي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =1856
                    Width =2665
                    Height =300
                    ColumnOrder =6
                    TabIndex =5
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9059
                            Top =1860
                            Width =735
                            Height =240
                            Name ="Label17"
                            Caption ="كد پستي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =7596
                    Width =2665
                    Height =300
                    ColumnWidth =1575
                    ColumnOrder =20
                    TabIndex =22
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8997
                            Top =7596
                            Width =660
                            Height =240
                            Name ="Label18"
                            Caption ="استان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =7937
                    Width =2665
                    Height =300
                    ColumnOrder =21
                    TabIndex =23
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8997
                            Top =7935
                            Width =930
                            Height =240
                            Name ="Label19"
                            Caption ="شهرستان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =2955
                    Width =2665
                    Height =300
                    ColumnOrder =7
                    TabIndex =8
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9029
                            Top =2955
                            Width =795
                            Height =240
                            Name ="Label20"
                            Caption ="كد ملي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =6237
                    Top =5568
                    Width =2665
                    Height =315
                    ColumnWidth =1785
                    ColumnOrder =13
                    TabIndex =14
                    ForeColor =8388608
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                        "TKNAME;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8998
                            Top =5568
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label113"
                            Caption ="نوع مشتري :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =5124
                    Width =2665
                    Height =300
                    ColumnOrder =12
                    TabIndex =13
                    Name ="CODE_E"
                    ControlSource ="CODE_E"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9023
                            Top =5123
                            Width =870
                            Height =240
                            Name ="Label22"
                            Caption ="ساير"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6237
                    Top =5987
                    Width =2665
                    Height =300
                    ColumnOrder =14
                    TabIndex =15
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9043
                            Top =5973
                            Width =2175
                            Height =240
                            Name ="Label24"
                            Caption ="*موبايل جهت ارسال پيامك:"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =877
                    Top =7107
                    Width =1588
                    Height =340
                    TabIndex =16
                    Name ="Command25"
                    Caption ="ذخيره و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =567
                    Top =6378
                    Width =8335
                    Height =317
                    ColumnWidth =7230
                    ColumnOrder =15
                    TabIndex =17
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.ROUTE_NAME + N' - ' + CUST_HESAB.NAME"
                        " + N' - ' + CUST_HESAB.hes AS Expr1 FROM Visit_route INNER JOIN CUST_HESAB ON Vi"
                        "sit_route.HES = CUST_HESAB.hes WHERE (Visit_route.RACTIVE = 1)"
                    ColumnWidths ="0"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9000
                            Top =6378
                            Width =1815
                            Height =330
                            ForeColor =8388608
                            Name ="Label26"
                            Caption ="مسير ويزيت"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3888
                    Top =340
                    Width =411
                    Height =300
                    ColumnOrder =16
                    TabIndex =18
                    Name ="HES"
                    ControlSource ="HES"
                    DefaultValue ="1"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =566
                    Width =411
                    Height =300
                    ColumnOrder =17
                    TabIndex =19
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    DefaultValue ="=[Forms]![Baseknow]![BEDEHKAR]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =6829
                    Width =2665
                    Height =300
                    ColumnOrder =18
                    TabIndex =20
                    Name ="Latitude"
                    ControlSource ="Latitude"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9030
                            Top =6828
                            Width =1110
                            Height =240
                            Name ="Label32"
                            Caption ="طول جغرافيايي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6237
                    Top =7250
                    Width =2665
                    Height =300
                    ColumnOrder =19
                    TabIndex =21
                    Name ="Longitude"
                    ControlSource ="Longitude"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9030
                            Top =7248
                            Width =1200
                            Height =240
                            Name ="Label34"
                            Caption ="عرض جغرافيايي"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =4125
                    Top =6828
                    Width =1395
                    Height =240
                    ForeColor =255
                    Name ="Label35"
                    Caption =" مثال :  31.844177"
                End
                Begin Label
                    OverlapFlags =85
                    Left =4125
                    Top =7188
                    Width =1395
                    Height =240
                    ForeColor =255
                    Name ="Label36"
                    Caption =" مثال : 54.384824"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6237
                    Top =2205
                    Width =2665
                    Height =300
                    ColumnOrder =9
                    TabIndex =6
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9060
                            Top =2210
                            Width =705
                            Height =240
                            Name ="Label38"
                            Caption ="کد استان"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6237
                    Top =2606
                    Width =2665
                    Height =300
                    ColumnOrder =10
                    TabIndex =7
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9038
                            Top =2600
                            Width =990
                            Height =240
                            Name ="Label40"
                            Caption ="کد شهرستان"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =680
                    Top =2097
                    Width =1589
                    Height =465
                    FontWeight =700
                    TabIndex =24
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =675
                    Top =2715
                    Width =1588
                    Height =465
                    TabIndex =25
                    Name ="Command41"
                    Caption ="اطلاعات بيشتر"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =6255
                    Top =8325
                    Width =2635
                    Height =300
                    TabIndex =26
                    ForeColor =8388608
                    Name ="tob"
                    ControlSource ="tob"
                    RowSourceType ="Value List"
                    RowSource ="1;\"حقيقي\";2;\"حقوقي\""
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9016
                            Top =8325
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label43"
                            Caption ="شخصيت :"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7426
                    Top =113
                    Width =685
                    Height =300
                    TabIndex =27
                    BackColor =13434828
                    Name ="TNUMBER2"
                    ControlSource ="TNUMBER2"
                    DefaultValue ="0"
                    Tag ="1374"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =8263
                            Top =113
                            Width =225
                            Height =285
                            Name ="Label45"
                            Caption ="كد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6519
                    Top =113
                    Width =685
                    Height =300
                    TabIndex =28
                    BackColor =16777164
                    Name ="TNUMBER3"
                    ControlSource ="TNUMBER3"
                    DefaultValue ="0"
                    Tag ="1374"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =7356
                            Top =113
                            Width =225
                            Height =285
                            Name ="Label47"
                            Caption ="كد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5839
                    Top =113
                    Width =685
                    Height =300
                    TabIndex =29
                    BackColor =16764057
                    Name ="TNUMBER4"
                    ControlSource ="TNUMBER4"
                    DefaultValue ="0"
                    Tag ="1374"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =6676
                            Top =113
                            Width =225
                            Height =285
                            Name ="Label49"
                            Caption ="كد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FCODE_CUSTOMER.cls"
