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
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =1
    ViewsAllowed =2
    TabularCharSet =178
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4616
    RowHeight =270
    ItemSuffix =24
    Left =6900
    Top =1470
    Right =12135
    Bottom =5490
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100
    RecSrcDt = Begin
        0xac71719610e7e540
    End
    RecordSource ="SELECT TNUMBER, NAME AS PERNAME, NUMBER, N_KOL FROM TDETA_HES WHERE (NUMBER = 1)"
        " AND (N_KOL = 213)"
    Caption ="."
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
            Height =1020
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =227
                    Width =2490
                    ColumnWidth =3660
                    ColumnOrder =5
                    Name ="PERNAME"
                    ControlSource ="PERNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3210
                            Top =225
                            Width =1380
                            Height =240
                            Name ="Label5"
                            Caption ="نام شخص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =426
                    Top =567
                    ColumnWidth =735
                    ColumnOrder =4
                    TabIndex =1
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2733
                            Top =567
                            Width =240
                            Height =240
                            Name ="Label17"
                            Caption ="كد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SERCHP.cls"
