Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =10397
    DatasheetFontHeight =10
    ItemSuffix =210
    Left =195
    Top =540
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x631a115533ede540
    End
    RecordSource ="SELECT SALA_DTL.EMZA AS emza1, SALA_DTL_1.EMZA AS emza2, SALA_DTL_2.EMZA AS emza"
        "3, PGET_HED.sgn1usid, PGET_HED.sgn2usid, PGET_HED.sgn3usid, PGET_HED.ID, sanadpa"
        "r_sub.MOLAH, sanadpar_sub.DATE_S, sanadpar_sub.SHOBEH, sanadpar_sub.mabld, sanad"
        "par_sub.N_HESAB, sanadpar_sub.DATE, sanadpar_sub.RADIF, sanadpar_sub.NO_AM, sana"
        "dpar_sub.NAHVA, sanadpar_sub.FHES_K, sanadpar_sub.FHES_M, sanadpar_sub.FHES_T, s"
        "anadpar_sub.THES_K, sanadpar_sub.THES_M, sanadpar_sub.THES_T, sanadpar_sub.SHARH"
        ", sanadpar_sub.MABL, sanadpar_sub.N_SERI, sanadpar_sub.BANK, sanadpar_sub.IDH, s"
        "anadpar_sub.FHES, sanadpar_sub.THES, sanadpar_sub.ARZD, sanadpar_sub.FHES_T2, sa"
        "nadpar_sub.THES_T2, sanadpar_sub.FHES_T3, sanadpar_sub.THES_T3, sanadpar_sub.FHE"
        "S_T4, sanadpar_sub.THES_T4, sanadpar_sub.hes, sanadpar_sub.NAME, sanadpar_sub.NA"
        "MES, sanadpar_sub.IDK, PGET_HED.SGN1, PGET_HED.SGN2, PGET_HED.SGN3 FROM SALA_DTL"
        " RIGHT OUTER JOIN SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN sanadpar_sub RIGHT OUTER "
        "JOIN PGET_HED ON sanadpar_sub.ID = PGET_HED.ID LEFT OUTER JOIN SALA_DTL SALA_DTL"
        "_2 ON PGET_HED.sgn3usid = SALA_DTL_2.IDD ON SALA_DTL_1.IDD = PGET_HED.sgn2usid O"
        "N SALA_DTL.IDD = PGET_HED.sgn1usid"
    Caption ="سند پرداخت"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="THES"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin BreakLevel
            ControlSource ="DATE_S"
        End
        Begin PageHeader
            Height =2622
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2266
                    Top =495
                    Width =5433
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4365
                            Width =1455
                            Height =600
                            FontSize =16
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label88"
                            Caption ="سند پرداخت"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =114
                    Top =450
                    Width =966
                    Height =390
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE"
                    ControlSource ="=[Forms]![pget_hed]![date]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1066
                            Top =450
                            Width =600
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =119
                    Top =56
                    Width =951
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="IDK"
                    ControlSource ="=[Forms]![pget_hed]![IDK]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1091
                            Top =56
                            Width =615
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Left =1
                    Top =960
                    Width =10335
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1728
                    Top =44
                    Width =411
                    Height =435
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="NAMEA"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1728
                    Top =496
                    Width =381
                    Height =435
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9766
                    Top =2112
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =8366
                    Top =2112
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =6872
                    Top =2112
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =5686
                    Top =2112
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =4213
                    Top =2112
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1741
                    Top =2112
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =346
                    Top =2157
                    Width =1215
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2386
                    Top =2160
                    Width =1395
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="شعبه"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4306
                    Top =2160
                    Width =1335
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="نام بانك"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6931
                    Top =2160
                    Width =1275
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="جاري"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5731
                    Top =2160
                    Width =975
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="تاريخ چك"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8386
                    Top =2154
                    Width =1380
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شماره چك"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9772
                    Top =2153
                    Width =540
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Top =1000
                    Width =10290
                    Height =1035
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =1
                    Top =2100
                    Width =10322
                    Height =510
                    Name ="Box196"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7847
                    Top =584
                    Width =2411
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2355
                    Top =165
                    Width =966
                    Height =390
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="DATE11"
                    ControlSource ="=[Forms]![pget_hed]![date]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =3307
                            Top =165
                            Width =600
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label209"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9781
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8385
                    Top =16
                    Width =1374
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="N_SERI"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6900
                    Top =15
                    Width =1419
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="N_HESAB"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4260
                    Top =16
                    Width =1389
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1771
                    Top =16
                    Width =2409
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =16
                    Width =1734
                    Height =405
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5736
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Top =450
                    Width =10335
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10305
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9765
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =8365
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =6871
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =5685
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =4212
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1740
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =3
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            ForceNewPage =2
            Height =2277
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =45
                    Width =1696
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =2
                    Left =8670
                    Top =30
                    Width =1680
                    Height =465
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label197"
                    Caption ="جمع كل به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    Left =2115
                    Top =30
                    Width =6525
                    Height =465
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR2"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BorderWidth =2
                    Width =10322
                    Height =510
                    Name ="Box201"
                End
                Begin Line
                    BorderWidth =1
                    Left =1740
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line204"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =6857
                    Top =576
                    Width =1701
                    Height =1701
                    Name ="EMZA1"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =3452
                    Top =561
                    Width =1701
                    Height =1701
                    TabIndex =1
                    Name ="EMZA2"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =17
                    Top =557
                    Width =1701
                    Height =1701
                    TabIndex =2
                    Name ="EMZA3"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8567
                    Top =996
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5145
                    Top =996
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1727
                    Top =996
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8564
                    Top =663
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5147
                    Top =663
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1712
                    Top =663
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5969
                    Top =1172
                    TabIndex =3
                    Name ="sgn1usid"
                    ControlSource ="sgn1usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4268
                            Top =1172
                            Width =930
                            Height =255
                            Name ="Label99"
                            Caption ="sgn1usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5969
                    Top =1512
                    TabIndex =4
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4268
                            Top =1512
                            Width =930
                            Height =255
                            Name ="Label100"
                            Caption ="sgn2usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5969
                    Top =1853
                    TabIndex =5
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =4268
                            Top =1853
                            Width =930
                            Height =255
                            Name ="Label101"
                            Caption ="sgn3usid:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1727
                    Top =1151
                    TabIndex =6
                    Name ="SGN1"
                    ControlSource ="SGN1"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1727
                    Top =1470
                    TabIndex =7
                    Name ="SGN2"
                    ControlSource ="SGN2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1727
                    Top =1789
                    TabIndex =8
                    Name ="SGN3"
                    ControlSource ="SGN3"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "SANAD_PARDAKHT.cls"
