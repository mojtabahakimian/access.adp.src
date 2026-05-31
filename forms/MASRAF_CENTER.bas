Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7823
    RowHeight =300
    ItemSuffix =23
    Left =4410
    Top =1980
    Right =11985
    Bottom =4470
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc04d31caefa0e240
    End
    RecordSource ="SELECT HEAD_MANF.* FROM HEAD_MANF WHERE ((Not (HEAD_MANF.N_KOL) Is Null)) OR ((N"
        "ot (HEAD_MANF.NAMES) Is Null));"
    Caption ="مراكز مصرف"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =2494
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
                    Left =4881
                    Top =90
                    Width =1401
                    Height =315
                    Name ="FNUMB"
                    ControlSource ="FNUMB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6408
                            Top =90
                            Width =1245
                            Height =300
                            Name ="Label1"
                            Caption ="كد مركز مصرف"
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
                    Left =2616
                    Top =453
                    Width =3681
                    Height =315
                    ColumnWidth =3060
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6408
                            Top =453
                            Width =1350
                            Height =300
                            Name ="Label19"
                            Caption ="نام  مركز مصرف"
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
                    ColumnCount =3
                    Left =171
                    Top =1575
                    Width =6126
                    Height =315
                    ColumnWidth =3525
                    TabIndex =4
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;2835;567"
                    OnGotFocus ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6408
                            Top =1575
                            Width =1110
                            Height =300
                            Name ="Label22"
                            Caption ="معين تفصيلي"
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
                    ColumnCount =3
                    Left =171
                    Top =1200
                    Width =6126
                    Height =315
                    ColumnWidth =2250
                    TabIndex =3
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6408
                            Top =1200
                            Width =1035
                            Height =300
                            Name ="Label21"
                            Caption ="معين"
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
                    ColumnCount =2
                    Left =171
                    Top =825
                    Width =6126
                    Height =315
                    ColumnWidth =2205
                    TabIndex =2
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6408
                            Top =825
                            Width =810
                            Height =300
                            Name ="Label20"
                            Caption ="حساب كل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MASRAF_CENTER.cls"
