Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
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
    Width =4605
    RowHeight =270
    ItemSuffix =31
    Left =6825
    Top =3510
    Right =21945
    Bottom =8895
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Tag ="SELECT     TDETA_HES.TNUMBER, TDETA_HES.NAME AS NAME, TDETA_HES.NUMBER, TDETA_HE"
        "S.N_KOL, DETA_HES.NAME AS NKOL, \015\012                      TOTA_HES.NAME AS N"
        "MOIN\015\012FROM         TOTA_HES INNER JOIN\015\012                      (DETA_"
        "HES INNER JOIN\015\012                      TDETA_HES ON (DETA_HES.NUMBER = TDET"
        "A_HES.NUMBER) AND (DETA_HES.N_KOL = TDETA_HES.N_KOL)) ON \015\012               "
        "       TOTA_HES.NUMBER = DETA_HES.N_KOL\015\012WHERE     (((TDETA_HES.CODE_E) LI"
        "KE IIf([Forms] ! [Baseknow] ! [PERSON] = 2, [0], [*]) OR\015\012                "
        "      (TDETA_HES.CODE_E) IS NULL))\015\012ORDER BY TDETA_HES.NAME;"
    RecSrcDt = Begin
        0x8056d532615be640
    End
    RecordSource ="SELECT TNUMBER, NAME, NUMBER, N_KOL, NMOIN, NKOL, ADDRESS, tNUMBER2, TNAME, CODE"
        "_E FROM cust_hesab_dtl WHERE     ((( NAME LIKE N'%مادر%' or  NAME LIKE N'%مادر%'"
        " OR NAME LIKE N'%مادر%' or  NAME LIKE  N'%مادر%') or (TNAME LIKE N'%مادر%' or TN"
        "AME LIKE N'%مادر%')))  ORDER BY NAME, TNAME "
    Caption ="."
    OnCurrent ="[Event Procedure]"
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
        Begin Section
            Height =6873
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
                    Top =737
                    Width =2490
                    ColumnWidth =3090
                    ColumnOrder =5
                    Name ="NAM"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =735
                            Width =1380
                            Height =240
                            Name ="Label5"
                            Caption ="نام حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =1360
                    ColumnWidth =825
                    ColumnOrder =4
                    TabIndex =1
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2533
                            Top =1360
                            Width =1095
                            Height =240
                            Name ="Label17"
                            Caption =" تفضيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =2040
                    ColumnWidth =900
                    ColumnOrder =2
                    TabIndex =2
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2533
                            Top =2040
                            Width =1095
                            Height =240
                            Name ="Label20"
                            Caption ="معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =2381
                    ColumnWidth =615
                    ColumnOrder =0
                    TabIndex =3
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2788
                            Top =2381
                            Width =840
                            Height =240
                            Name ="Label21"
                            Caption =" كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =396
                    Top =2834
                    ColumnWidth =2070
                    ColumnOrder =3
                    TabIndex =4
                    Name ="NMOIN"
                    ControlSource ="NMOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2553
                            Top =2834
                            Width =1245
                            Height =240
                            Name ="Label22"
                            Caption ="نام معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =283
                    Top =3458
                    ColumnWidth =1950
                    ColumnOrder =1
                    TabIndex =5
                    Name ="NKOL"
                    ControlSource ="NKOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2845
                            Top =3458
                            Width =840
                            Height =240
                            Name ="Label23"
                            Caption ="نام كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnWidth =6780
                    ColumnOrder =8
                    TabIndex =6
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2562
                            Width =840
                            Height =240
                            Name ="Label25"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =1720
                    ColumnWidth =1350
                    ColumnOrder =6
                    TabIndex =7
                    Name ="TNUMBER2"
                    ControlSource ="TNUMBER2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2535
                            Top =1725
                            Width =1305
                            Height =240
                            Name ="Label28"
                            Caption =" تفضيلي هاي بعد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1099
                    Width =2490
                    ColumnWidth =2865
                    ColumnOrder =7
                    TabIndex =8
                    Name ="TNAME"
                    ControlSource ="TNAME"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3180
                            Top =1095
                            Width =1425
                            Height =405
                            Name ="Label30"
                            Caption ="نام حساب تفصيليهاي بعد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SERSND.cls"
