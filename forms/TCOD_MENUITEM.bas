Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9077
    RowHeight =345
    ItemSuffix =11
    Left =3900
    Top =1050
    Right =12495
    Bottom =6315
    HelpContextId =45
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8479a2ce315be440
    End
    RecordSource ="TCODE_MENUITEM"
    Caption ="تعريف گروه بندي براي ويزيتور"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
            Height =5272
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5458
                    Top =113
                    Width =1416
                    Height =300
                    ColumnWidth =765
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6945
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
                    Left =3364
                    Top =453
                    Width =3510
                    Height =300
                    ColumnWidth =4665
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6955
                            Top =450
                            Width =600
                            Height =285
                            Name ="Label3"
                            Caption ="نام گروه"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1804
                    Top =876
                    Width =5070
                    Height =330
                    ColumnWidth =3825
                    TabIndex =2
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6962
                            Top =870
                            Width =2115
                            Height =240
                            Name ="Label5"
                            Caption ="انباري كه كالا از آن كم ميشود "
                        End
                    End
                End
                Begin Image
                    SizeMode =3
                    PictureType =1
                    Left =2324
                    Top =1814
                    Width =4331
                    Height =2669
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1247
                    Top =4694
                    Width =5676
                    Height =315
                    TabIndex =3
                    Name ="Text9"
                    ControlSource ="=[Forms]![BASEKNOW]![BACKPATH] & \"\\grp\\\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6984
                            Top =4702
                            Width =1335
                            Height =300
                            Name ="Label10"
                            Caption ="مسير فايل تصاوير:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TCOD_MENUITEM.cls"
