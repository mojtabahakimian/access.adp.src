Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =1
    ViewsAllowed =2
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5730
    RowHeight =315
    ItemSuffix =21
    Left =3780
    Top =195
    Right =10545
    Bottom =4980
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x071ca6246ddde240
    End
    RecordSource ="SELECT HEAD_MANF.FNUMB, STUF_DEF.CODE, STUF_DEF.NAME FROM STUF_DEF INNER JOIN HE"
        "AD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE GROUP BY STUF_DEF.CODE, STUF_DEF.NAME,"
        " HEAD_MANF.FNUMB"
    Caption ="."
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    AllowFormView =0
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
            Height =2097
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    ColumnWidth =855
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5114
                            Width =555
                            Height =240
                            Name ="Label16"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =340
                    ColumnWidth =4995
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5114
                            Top =340
                            Width =555
                            Height =240
                            Name ="Label17"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =680
                    ColumnWidth =990
                    TabIndex =2
                    Name ="FNUMB"
                    ControlSource ="FNUMB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4485
                            Top =685
                            Width =840
                            Height =240
                            Name ="Label18"
                            Caption ="شماره فرمول"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SERCHJOB.cls"
