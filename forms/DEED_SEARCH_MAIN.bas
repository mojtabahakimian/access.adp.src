Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11055
    RowHeight =375
    ItemSuffix =103
    Left =3765
    Top =585
    Right =14820
    Bottom =7740
    DatasheetGridlinesColor =12632256
    MaxRecords =100
    ShortcutMenuBar ="SERVERFILTER"
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0x1e8be3741ceee240
    End
    Caption ="جستجو در شرح سند"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin FormHeader
            Height =690
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =4884
                    Top =15
                    Width =2100
                    Height =675
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label30"
                    Caption ="جستجو در اسناد"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =2
                    Left =4839
                    Width =2100
                    Height =675
                    FontSize =18
                    FontWeight =700
                    Name ="Label31"
                    Caption ="جستجو در اسناد"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =6474
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =480
                    Width =2436
                    Height =313
                    ColumnWidth =990
                    ColumnOrder =0
                    TabIndex =1
                    BackColor =16777164
                    Name ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9139
                            Top =480
                            Width =1320
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =1920
                    Width =2436
                    Height =313
                    ColumnWidth =5010
                    ColumnOrder =5
                    TabIndex =5
                    BackColor =16777164
                    Name ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =1920
                            Width =615
                            Height =240
                            Name ="Label9"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =2280
                    Width =2436
                    Height =313
                    ColumnWidth =1560
                    ColumnOrder =6
                    TabIndex =6
                    BackColor =16777164
                    Name ="BED"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =2280
                            Width =720
                            Height =240
                            Name ="Label11"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =2640
                    Width =2436
                    Height =313
                    ColumnWidth =1545
                    ColumnOrder =7
                    TabIndex =7
                    BackColor =16777164
                    Name ="BES"
                    Format ="#,##0.00;-#,##0.00"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =2640
                            Width =840
                            Height =240
                            Name ="Label13"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =3000
                    Width =2436
                    Height =313
                    ColumnWidth =990
                    ColumnOrder =9
                    TabIndex =8
                    BackColor =16777164
                    Name ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =3000
                            Width =1155
                            Height =240
                            Name ="Label15"
                            Caption ="شماره سري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =3720
                    Width =2436
                    Height =313
                    ColumnWidth =1125
                    ColumnOrder =11
                    TabIndex =10
                    BackColor =16777164
                    Name ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =3720
                            Width =1155
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =4443
                    Width =2436
                    Height =313
                    ColumnWidth =5250
                    ColumnOrder =8
                    TabIndex =12
                    BackColor =16777164
                    Name ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =4440
                            Width =1125
                            Height =240
                            Name ="Label22"
                            Caption ="شرح حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =120
                    Width =2436
                    Height =313
                    ColumnWidth =900
                    ColumnOrder =1
                    BackColor =16777164
                    Name ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =120
                            Width =945
                            Height =240
                            Name ="Label27"
                            Caption ="تاريخ سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =840
                    Width =2436
                    Height =313
                    ColumnWidth =525
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =16777164
                    Name ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =840
                            Width =510
                            Height =240
                            Name ="Label5"
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
                    Left =6523
                    Top =1195
                    Width =2436
                    Height =313
                    ColumnWidth =585
                    ColumnOrder =3
                    TabIndex =3
                    BackColor =16777164
                    Name ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =1195
                            Width =675
                            Height =240
                            Name ="Label7"
                            Caption =" معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6523
                    Top =1560
                    Width =2436
                    Height =313
                    ColumnWidth =825
                    ColumnOrder =4
                    TabIndex =4
                    BackColor =16777164
                    Name ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =1567
                            Width =870
                            Height =240
                            Name ="Label29"
                            Caption =" تفصيلي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5489
                    Top =113
                    Width =816
                    Height =314
                    TabIndex =13
                    BackColor =13434828
                    Name ="DATE_SB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =480
                    Width =816
                    Height =314
                    TabIndex =14
                    BackColor =13434828
                    Name ="N_SB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =841
                    Width =816
                    Height =314
                    TabIndex =15
                    BackColor =13434828
                    Name ="HES_KB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =1195
                    Width =816
                    Height =314
                    TabIndex =16
                    BackColor =13434828
                    Name ="HES_MB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =1561
                    Width =816
                    Height =314
                    TabIndex =17
                    BackColor =13434828
                    Name ="TNUMBERB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =1920
                    Width =816
                    Height =314
                    TabIndex =18
                    BackColor =13434828
                    Name ="SHARHB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =2281
                    Width =816
                    Height =314
                    TabIndex =19
                    BackColor =13434828
                    Name ="BEDB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =2641
                    Width =816
                    Height =314
                    TabIndex =20
                    BackColor =13434828
                    Name ="BESB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =3001
                    Width =816
                    Height =314
                    TabIndex =21
                    BackColor =13434828
                    Name ="N_SERIB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =3361
                    Width =816
                    Height =314
                    TabIndex =22
                    BackColor =13434828
                    Name ="BANKB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =3720
                    Width =816
                    Height =314
                    TabIndex =23
                    BackColor =13434828
                    Name ="NUMBERB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =4081
                    Width =816
                    Height =314
                    TabIndex =24
                    BackColor =13434828
                    Name ="TAGB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5490
                    Top =4443
                    Width =816
                    Height =314
                    TabIndex =25
                    BackColor =13434828
                    Name ="hesB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6523
                    Top =3360
                    Width =2436
                    Height =315
                    ColumnWidth =465
                    ColumnOrder =10
                    TabIndex =9
                    BackColor =16777164
                    Name ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_BANKS ORDER BY NAMES"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =3360
                            Width =540
                            Height =240
                            Name ="Label17"
                            Caption ="بانك"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6523
                    Top =4080
                    Width =2436
                    Height =315
                    ColumnWidth =645
                    ColumnOrder =12
                    TabIndex =11
                    BackColor =16777164
                    Name ="HTAG"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, BARGAH FROM TAGCOD ORDER BY BARGAH"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9137
                            Top =4080
                            Width =825
                            Height =240
                            Name ="Label21"
                            Caption ="نوع برگه"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3855
                    Top =5550
                    Width =2158
                    Height =402
                    TabIndex =26
                    Name ="Command71"
                    Caption ="اجراي گزارش"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =1695
                    Top =5555
                    Width =2158
                    Height =402
                    TabIndex =27
                    Name ="Command72"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =480
                    Width =2436
                    Height =313
                    TabIndex =28
                    BackColor =16777164
                    Name ="N_ST"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3981
                            Top =480
                            Width =1320
                            Height =240
                            Name ="Label76"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =2280
                    Width =2436
                    Height =313
                    TabIndex =29
                    BackColor =16777164
                    Name ="BEDT"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =2280
                            Width =720
                            Height =240
                            Name ="Label80"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =2640
                    Width =2436
                    Height =313
                    TabIndex =30
                    BackColor =16777164
                    Name ="BEST"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =2640
                            Width =840
                            Height =240
                            Name ="Label82"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =3720
                    Width =2436
                    Height =313
                    TabIndex =31
                    BackColor =16777164
                    Name ="NUMBERT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =3720
                            Width =1155
                            Height =240
                            Name ="Label86"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =120
                    Width =2436
                    Height =313
                    TabIndex =32
                    BackColor =16777164
                    Name ="DATE_ST"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =120
                            Width =945
                            Height =240
                            Name ="Label90"
                            Caption ="تاريخ سند"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =840
                    Width =2436
                    Height =313
                    TabIndex =33
                    BackColor =16777164
                    Name ="HES_KT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =840
                            Width =510
                            Height =240
                            Name ="Label92"
                            Caption =" كل"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =1195
                    Width =2436
                    Height =313
                    TabIndex =34
                    BackColor =16777164
                    Name ="HES_MT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =1195
                            Width =675
                            Height =240
                            Name ="Label94"
                            Caption =" معين"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1365
                    Top =1560
                    Width =2436
                    Height =313
                    TabIndex =35
                    BackColor =16777164
                    Name ="TNUMBERT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3979
                            Top =1567
                            Width =870
                            Height =240
                            Name ="Label96"
                            Caption =" تفصيلي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7905
                    Top =4815
                    Width =1041
                    Height =389
                    FontSize =12
                    FontWeight =700
                    TabIndex =36
                    Name ="ANDOR"
                    RowSourceType ="Value List"
                    RowSource ="و;يا"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"يا\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9102
                            Top =4865
                            Width =1455
                            Height =240
                            Name ="Label102"
                            Caption ="شروط بعدي بوسيله"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "DEED_SEARCH_MAIN.cls"
