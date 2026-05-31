Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7237
    RowHeight =285
    ItemSuffix =10
    Left =1485
    Top =795
    Right =10515
    Bottom =4695
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x04f4c1cdbc9ee240
    End
    RecordSource ="ROOM"
    Caption ="تعريف اتاقها"
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
        Begin Section
            Height =1834
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4983
                    Top =113
                    Width =1026
                    Height =255
                    ColumnWidth =1050
                    FontWeight =700
                    Name ="ROOM_NUM"
                    ControlSource ="ROOM_NUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =6125
                            Top =113
                            Width =975
                            Height =240
                            FontWeight =700
                            Name ="Label1"
                            Caption ="شماره اتاق"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4983
                    Top =453
                    Width =1026
                    Height =255
                    ColumnWidth =705
                    FontWeight =700
                    TabIndex =1
                    Name ="RFLAG"
                    ControlSource ="RFLAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =6125
                            Top =453
                            Width =630
                            Height =240
                            FontWeight =700
                            Name ="Label3"
                            Caption ="وضعيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4983
                    Top =793
                    Width =1026
                    Height =255
                    ColumnWidth =930
                    FontWeight =700
                    TabIndex =2
                    Name ="TAKHT"
                    ControlSource ="TAKHT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =6125
                            Top =793
                            Width =855
                            Height =240
                            FontWeight =700
                            Name ="Label5"
                            Caption ="تعداد تخت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    Width =5865
                    Height =255
                    ColumnWidth =6870
                    FontWeight =700
                    TabIndex =3
                    Name ="VASAYEL"
                    ControlSource ="VASAYEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =6125
                            Top =1133
                            Width =990
                            Height =240
                            FontWeight =700
                            Name ="Label7"
                            Caption ="وسايل اتاق"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1474
                    Width =5865
                    Height =255
                    ColumnWidth =1815
                    FontWeight =700
                    TabIndex =4
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =6125
                            Top =1474
                            Width =765
                            Height =240
                            FontWeight =700
                            Name ="Label9"
                            Caption ="توضيحات"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ROOM.cls"
