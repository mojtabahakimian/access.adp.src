Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =60
    GridY =60
    Width =14494
    DatasheetFontHeight =10
    ItemSuffix =269
    Left =240
    Right =14745
    Bottom =7980
    HelpContextId =48
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x1482b83c990be440
    End
    RecordSource ="TR_SAZMAN"
    Caption ="."
    BeforeUpdate ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            SpecialEffect =1
            Height =13483
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9507
                    Top =5339
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12628
                            Top =5336
                            Width =900
                            Height =240
                            Name ="Label148"
                            Caption ="تاريخ اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9507
                    Top =6394
                    Width =3606
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13196
                            Top =6396
                            Width =1005
                            Height =240
                            Name ="Label149"
                            Caption ="ساعت اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9507
                    Top =6867
                    Width =3111
                    Height =345
                    TabIndex =2
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12637
                            Top =6912
                            Width =1380
                            Height =285
                            Name ="Label229"
                            Caption ="كاربر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9507
                    Top =5878
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12523
                            Top =5873
                            Width =1755
                            Height =240
                            Name ="Label230"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =9507
                    Top =7343
                    Width =3339
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12875
                            Top =7395
                            Width =300
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =60
                    Width =997
                    Height =315
                    TabIndex =5
                    Name ="SANDOGH"
                    ControlSource ="SANDOGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12971
                            Top =60
                            Width =690
                            Height =300
                            Name ="Label117"
                            Caption =" صندوق:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =400
                    Width =997
                    Height =315
                    TabIndex =6
                    Name ="BANKHA"
                    ControlSource ="BANKHA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12971
                            Top =400
                            Width =615
                            Height =300
                            Name ="Label118"
                            Caption =" بانكها:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =740
                    Width =997
                    Height =255
                    TabIndex =7
                    Name ="BESTANKAR"
                    ControlSource ="BESTANKAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12971
                            Top =740
                            Width =960
                            Height =300
                            Name ="Label119"
                            Caption =" بستانكاران:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =1081
                    Width =997
                    Height =315
                    TabIndex =8
                    Name ="BEDEHKAR"
                    ControlSource ="BEDEHKAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12971
                            Top =1081
                            Width =870
                            Height =300
                            Name ="Label120"
                            Caption =" بدهكاران:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =1421
                    Width =997
                    Height =315
                    TabIndex =9
                    Name ="KHARID"
                    ControlSource ="KHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =12975
                            Top =1425
                            Width =1035
                            Height =405
                            Name ="Label121"
                            Caption =" كنترل خريد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =1761
                    Width =997
                    Height =315
                    TabIndex =10
                    Name ="MKHARID"
                    ControlSource ="MKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =12971
                            Top =1761
                            Width =1185
                            Height =300
                            Name ="Label122"
                            Caption =" مرجوعي خريد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =2101
                    Width =997
                    Height =315
                    TabIndex =11
                    Name ="TKHARID"
                    ControlSource ="TKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12968
                            Top =2100
                            Width =1155
                            Height =300
                            Name ="Label123"
                            Caption =" تخفيفات خريد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =2441
                    Width =997
                    Height =315
                    TabIndex =12
                    Name ="HKHARID"
                    ControlSource ="HKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12968
                            Top =2445
                            Width =1350
                            Height =300
                            Name ="Label124"
                            Caption =" هزينه هاي خريد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =2767
                    Width =997
                    Height =315
                    TabIndex =13
                    Name ="FROSH"
                    ControlSource ="FROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12971
                            Top =2767
                            Width =630
                            Height =300
                            Name ="Label125"
                            Caption =" فروش:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =3107
                    Width =997
                    Height =315
                    TabIndex =14
                    Name ="MFROSH"
                    ControlSource ="MFROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12968
                            Top =3112
                            Width =1305
                            Height =300
                            Name ="Label126"
                            Caption =" مرجوعي فروش:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11878
                    Top =3457
                    Width =997
                    Height =315
                    TabIndex =15
                    Name ="TFROSH"
                    ControlSource ="TFROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12968
                            Top =3450
                            Width =1275
                            Height =300
                            Name ="Label127"
                            Caption =" تخفيفات فروش:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11858
                    Top =3840
                    Width =996
                    Height =315
                    TabIndex =16
                    Name ="HFROSH"
                    ControlSource ="HFROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12953
                            Top =3840
                            Width =1425
                            Height =300
                            Name ="Label128"
                            Caption =" هزينه‌هاي فروش:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =50
                    Width =996
                    Height =315
                    TabIndex =17
                    Name ="MOGODIA"
                    ControlSource ="MOGODIA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6841
                            Top =45
                            Width =1470
                            Height =300
                            Name ="Label129"
                            Caption ="انبار:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =420
                    Width =996
                    Height =315
                    TabIndex =18
                    Name ="DARAM"
                    ControlSource ="DARAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =420
                            Width =1410
                            Height =300
                            Name ="Label131"
                            Caption ="درآمد ها:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =760
                    Width =996
                    Height =315
                    TabIndex =19
                    Name ="HDARAM"
                    ControlSource ="HDARAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =800
                            Width =1110
                            Height =240
                            Name ="Label132"
                            Caption ="تخفيفات درآمد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =1125
                    Width =3486
                    Height =315
                    TabIndex =20
                    Name ="ADA"
                    ControlSource ="ADA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES ORDER BY NA"
                        "ME"
                    ColumnWidths ="1134"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =6840
                            Top =1125
                            Width =2265
                            Height =405
                            Name ="Label134"
                            Caption =" اسناد دريافتني تجاري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3240
                    Top =1465
                    Width =3501
                    Height =315
                    TabIndex =21
                    Name ="APA"
                    ControlSource ="APA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES ORDER BY NA"
                        "ME"
                    ColumnWidths ="1134"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =6841
                            Top =1459
                            Width =2250
                            Height =300
                            Name ="Label135"
                            Caption =" اسناد پرداختني تجاري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =2586
                    Width =996
                    Height =315
                    TabIndex =22
                    Name ="HAVALAH"
                    ControlSource ="HAVALAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =2585
                            Width =1410
                            Height =300
                            Name ="Label137"
                            Caption =" حواله هاي بانكي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =82
                    Width =2407
                    Height =315
                    TabIndex =23
                    Name ="Combo144"
                    ControlSource ="SANDOGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =422
                    Width =2407
                    Height =315
                    TabIndex =24
                    Name ="Combo146"
                    ControlSource ="BANKHA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =762
                    Width =2407
                    Height =315
                    TabIndex =25
                    Name ="Combo148"
                    ControlSource ="BESTANKAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =1103
                    Width =2407
                    Height =315
                    TabIndex =26
                    Name ="Combo150"
                    ControlSource ="BEDEHKAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =1443
                    Width =2407
                    Height =315
                    TabIndex =27
                    Name ="Combo152"
                    ControlSource ="KHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =1783
                    Width =2407
                    Height =315
                    TabIndex =28
                    Name ="Combo154"
                    ControlSource ="MKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =2123
                    Width =2407
                    Height =315
                    TabIndex =29
                    Name ="Combo156"
                    ControlSource ="TKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =2463
                    Width =2407
                    Height =315
                    TabIndex =30
                    Name ="Combo158"
                    ControlSource ="HKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =2789
                    Width =2407
                    Height =315
                    TabIndex =31
                    Name ="Combo160"
                    ControlSource ="FROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =3129
                    Width =2407
                    Height =315
                    TabIndex =32
                    Name ="Combo162"
                    ControlSource ="MFROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9353
                    Top =3479
                    Width =2407
                    Height =315
                    TabIndex =33
                    Name ="Combo164"
                    ControlSource ="TFROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9368
                    Top =3840
                    Width =2407
                    Height =315
                    TabIndex =34
                    Name ="Combo166"
                    ControlSource ="HFROSH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =50
                    Width =2407
                    Height =315
                    TabIndex =35
                    Name ="Combo167"
                    ControlSource ="MOGODIA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =421
                    Width =2407
                    Height =315
                    TabIndex =36
                    Name ="Combo169"
                    ControlSource ="DARAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =761
                    Width =2407
                    Height =315
                    TabIndex =37
                    Name ="Combo170"
                    ControlSource ="HDARAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =2580
                    Width =2407
                    Height =315
                    TabIndex =38
                    Name ="Combo175"
                    ControlSource ="HAVALAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =2936
                    Width =996
                    Height =315
                    TabIndex =39
                    Name ="HAZ_TOL"
                    ControlSource ="HAZ_TOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =2935
                            Width =1980
                            Height =300
                            Name ="Label182"
                            Caption =" كنترل مواد مصرفي توليد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =3276
                    Width =996
                    Height =315
                    TabIndex =40
                    Name ="PHAZ_TOL"
                    ControlSource ="PHAZ_TOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =3282
                            Width =2010
                            Height =300
                            Name ="Label184"
                            Caption ="جذب هزينه توليد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =2930
                    Width =2407
                    Height =315
                    TabIndex =41
                    Name ="HAZ_TOL2"
                    ControlSource ="HAZ_TOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =3270
                    Width =2407
                    Height =315
                    TabIndex =42
                    Name ="PHAZ_TOL11"
                    ControlSource ="PHAZ_TOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =4002
                    Width =996
                    Height =315
                    TabIndex =43
                    Name ="PPDAST"
                    ControlSource ="PPDAST"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =4008
                            Width =1590
                            Height =300
                            Name ="Label188"
                            Caption ="كنترل دستمزد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =3996
                    Width =2407
                    Height =315
                    TabIndex =44
                    Name ="PPDAST1"
                    ControlSource ="PPDAST"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =4377
                    Width =996
                    Height =315
                    TabIndex =45
                    Name ="PPSAR"
                    ControlSource ="PPSAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =4383
                            Width =1485
                            Height =300
                            Name ="Label191"
                            Caption ="كنترل سربار توليد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =4371
                    Width =2407
                    Height =315
                    TabIndex =46
                    Name ="PPSAR1"
                    ControlSource ="PPSAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =5110
                    Width =996
                    Height =315
                    TabIndex =47
                    Name ="GHEYMAT"
                    ControlSource ="GHEYMAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =5119
                            Width =2535
                            Height =300
                            Name ="Label194"
                            Caption ="قيمت تمام شده كالاي فروش رفته:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =5104
                    Width =2407
                    Height =315
                    TabIndex =48
                    Name ="GHEYMAT1"
                    ControlSource ="GHEYMAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5746
                    Top =5483
                    Width =996
                    Height =315
                    TabIndex =49
                    Name ="AMALKARD"
                    ControlSource ="AMALKARD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6841
                            Top =5492
                            Width =1590
                            Height =300
                            Name ="Label197"
                            Caption ="عملكرد :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3256
                    Top =5477
                    Width =2407
                    Height =315
                    TabIndex =50
                    Name ="AMALKARD1"
                    ControlSource ="AMALKARD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11859
                    Top =4206
                    Width =996
                    Height =315
                    TabIndex =51
                    Name ="PERSONEL"
                    ControlSource ="PERSONEL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12955
                            Top =4205
                            Width =585
                            Height =300
                            Name ="Label199"
                            Caption ="پرسنل:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11859
                    Top =4546
                    Width =996
                    Height =315
                    TabIndex =52
                    Name ="PERVAM"
                    ControlSource ="PERVAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12954
                            Top =4545
                            Width =885
                            Height =300
                            Name ="Label190"
                            Caption ="وام پرسنل:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9369
                    Top =4200
                    Width =2407
                    Height =315
                    TabIndex =53
                    Name ="PERSONEL1"
                    ControlSource ="PERSONEL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =9369
                    Top =4540
                    Width =2407
                    Height =315
                    TabIndex =54
                    Name ="PERVAM1"
                    ControlSource ="PERVAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5730
                    Top =4750
                    Width =996
                    Height =315
                    TabIndex =55
                    Name ="CONKAL"
                    ControlSource ="CONKAL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6825
                            Top =4759
                            Width =2175
                            Height =300
                            Name ="Label201"
                            Caption ="كنترل كالاي در جريان ساخت:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3240
                    Top =4744
                    Width =2407
                    Height =315
                    TabIndex =56
                    Name ="CONKAL2"
                    ControlSource ="CONKAL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =3642
                    Width =996
                    Height =315
                    TabIndex =57
                    Name ="PJHAZ_TOL1"
                    ControlSource ="PJHAZ_TOL1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =3651
                            Width =2355
                            Height =300
                            Name ="Label207"
                            Caption ="پاياپاي هزينه هاي توليد و خدمات:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =3636
                    Width =2407
                    Height =315
                    TabIndex =58
                    Name ="PJHAZ_TOL12"
                    ControlSource ="PJHAZ_TOL1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5745
                    Top =5862
                    Width =996
                    Height =315
                    TabIndex =59
                    Name ="PKHARID"
                    ControlSource ="PKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =5871
                            Width =1590
                            Height =300
                            Name ="Label210"
                            Caption ="پاياپاي كنترل خريد :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =5856
                    Width =2407
                    Height =315
                    TabIndex =60
                    Name ="PKHARID2"
                    ControlSource ="PKHARID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =6207
                    Width =3486
                    Height =315
                    TabIndex =61
                    Name ="HESMBAA"
                    ControlSource ="HESMBAA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES ORDER BY NA"
                        "ME"
                    ColumnWidths ="1134"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =6216
                            Width =1830
                            Height =285
                            Name ="Label213"
                            Caption ="ماليات بر ارزش افزوده :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =6642
                    Width =3486
                    Height =315
                    TabIndex =62
                    Name ="HPOR"
                    ControlSource ="HPOR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES ORDER BY NA"
                        "ME"
                    ColumnWidths ="1134"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6840
                            Top =6651
                            Width =1830
                            Height =285
                            Name ="Label217"
                            Caption ="پورسانت فروش :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3255
                    Top =1830
                    Width =3486
                    Height =315
                    TabIndex =63
                    Name ="ADV"
                    ControlSource ="ADV"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES ORDER BY NA"
                        "ME"
                    ColumnWidths ="1134"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =6840
                            Top =1830
                            Width =2265
                            Height =405
                            Name ="Label219"
                            Caption =" اسناد دريافتني غير تجاري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3240
                    Top =2170
                    Width =3501
                    Height =315
                    TabIndex =64
                    Name ="APV"
                    ControlSource ="APV"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES ORDER BY NA"
                        "ME"
                    ColumnWidths ="1134"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =6841
                            Top =2164
                            Width =2250
                            Height =300
                            Name ="Label221"
                            Caption =" اسناد پرداختني غير تجاري:"
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1065
                    Top =7155
                    Width =7158
                    Height =6210
                    FontWeight =700
                    TabIndex =65
                    Name ="TabCtl137"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =1200
                            Top =7560
                            Width =6885
                            Height =5670
                            Name ="Page138"
                            Caption ="عمومي"
                            Begin
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1760
                                    Top =7784
                                    Width =4207
                                    Height =300
                                    Name ="WIDTH_D"
                                    ControlSource ="WIDTH_D"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =6079
                                            Top =7789
                                            Width =1245
                                            Height =285
                                            Name ="Label37"
                                            Caption ="متن بالاي فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1760
                                    Top =8124
                                    Width =4207
                                    Height =300
                                    TabIndex =1
                                    Name ="HIGH_D"
                                    ControlSource ="HIGH_D"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =6079
                                            Top =8119
                                            Width =1185
                                            Height =285
                                            Name ="Label39"
                                            Caption ="متن پائين فاكتور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    Left =1760
                                    Top =8479
                                    Width =4207
                                    Height =300
                                    TabIndex =2
                                    Name ="BACKPATH"
                                    ControlSource ="BACKPATH"
                                    DefaultValue ="\"D:\\DENABACK\""
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =6079
                                            Top =8474
                                            Width =1470
                                            Height =240
                                            Name ="Label97"
                                            Caption ="مسير فايل پشتيبان:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1595
                                    Top =10223
                                    Width =5166
                                    Height =300
                                    TabIndex =3
                                    Name ="TFADDRESS"
                                    ControlSource ="TFADDRESS"
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextFontCharSet =0
                                            Left =6842
                                            Top =10223
                                            Width =615
                                            Height =285
                                            Name ="Label114"
                                            Caption ="آدرس:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =1490
                                    Top =10058
                                    Width =5908
                                    Height =1021
                                    Name ="Box116"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =223
                                    TextFontCharSet =0
                                    Left =6440
                                    Top =9773
                                    Width =885
                                    Height =255
                                    BackColor =-2147483633
                                    Name ="Label243"
                                    Caption ="آدرس ، تلفن"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4040
                                    Top =10583
                                    Width =2706
                                    Height =300
                                    TabIndex =4
                                    Name ="TFTEL"
                                    ControlSource ="TFTEL"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6872
                                            Top =10583
                                            Width =525
                                            Height =285
                                            Name ="Label115"
                                            Caption ="تلفن:"
                                        End
                                    End
                                End
                                Begin BoundObjectFrame
                                    OverlapFlags =215
                                    Left =1432
                                    Top =8926
                                    Width =6145
                                    Height =860
                                    TabIndex =5
                                    Name ="OLEBound157"
                                    ControlSource ="EMZA"
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =1
                                    Left =4128
                                    Top =11250
                                    Width =2706
                                    Height =300
                                    TabIndex =6
                                    Name ="SERVERNAM"
                                    ControlSource ="SERVERNAM"
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            Left =6955
                                            Top =11249
                                            Width =750
                                            Height =285
                                            Name ="Label313"
                                            Caption ="نام سرور:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1539
                                    Top =11704
                                    Width =4821
                                    Height =255
                                    TabIndex =7
                                    Name ="OPTIONSS"
                                    ControlSource ="OPTIONSS"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =1200
                            Top =7560
                            Width =6885
                            Height =5670
                            Name ="Page139"
                            Caption ="اسناد وحساب"
                            Begin
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =1385
                                    Top =7768
                                    Width =2717
                                    Height =1318
                                    Name ="Frame98"
                                    ControlSource ="SANAD"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2754
                                            Top =7654
                                            Width =675
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Label99"
                                            Caption ="سند زدن"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3752
                                            Top =8021
                                            OptionValue =1
                                            Name ="Option101"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =2275
                                                    Top =7964
                                                    Width =1425
                                                    Height =255
                                                    Name ="Label102"
                                                    Caption ="در هر زمان اتوماتيك"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3752
                                            Top =8361
                                            OptionValue =2
                                            Name ="Option103"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1495
                                                    Top =8304
                                                    Width =2205
                                                    Height =285
                                                    Name ="Label104"
                                                    Caption ="در زمان تائيد توسط حسابداري"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3737
                                            Top =8692
                                            OptionValue =3
                                            Name ="Option149"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =2140
                                                    Top =8645
                                                    Width =1560
                                                    Height =285
                                                    Name ="Label150"
                                                    Caption ="هيچكدام"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =4675
                                    Top =7781
                                    Width =2579
                                    Height =943
                                    Name ="Frame124"
                                    ControlSource ="PERSON"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =4780
                                            Top =7619
                                            Width =2385
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label244"
                                            Caption ="اشخاص مشخص شده  در حسابها"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6571
                                            Top =8019
                                            OptionValue =1
                                            Name ="Option126"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5192
                                                    Top =7962
                                                    Width =1410
                                                    Height =285
                                                    Name ="Label245"
                                                    Caption ="نمايش داده شوند"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6571
                                            Top =8349
                                            OptionValue =2
                                            Name ="Option128"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5177
                                                    Top =8302
                                                    Width =1425
                                                    Height =285
                                                    Name ="Label246"
                                                    Caption ="نمايش داده نشوند"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =4645
                                    Top =9461
                                    Width =2620
                                    Height =1273
                                    Name ="TKHF"
                                    ControlSource ="TKHF"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =5328
                                            Top =9299
                                            Width =1215
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label152"
                                            Caption ="تخفيفات فاكتورها"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6571
                                            Top =9699
                                            OptionValue =2
                                            Name ="Option153"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4907
                                                    Top =9663
                                                    Width =1695
                                                    Height =285
                                                    Name ="Label154"
                                                    Caption ="مصوب عمل شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6571
                                            Top =10029
                                            OptionValue =1
                                            Name ="Option155"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4645
                                                    Top =10004
                                                    Width =1950
                                                    Height =285
                                                    Name ="Label156"
                                                    Caption ="مبلغ كلي پشت فاكتور"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =6564
                                            Top =10375
                                            OptionValue =3
                                            Name ="Option300"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4795
                                                    Top =10349
                                                    Width =1800
                                                    Height =285
                                                    Name ="Label301"
                                                    Caption ="به ريز كالا عمل شود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    KeyboardLanguage =43
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4675
                                    Top =11416
                                    Width =621
                                    Height =255
                                    Name ="TFTPAGE"
                                    ControlSource ="TFTPAGE"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =5355
                                            Top =11415
                                            Width =2325
                                            Height =240
                                            Name ="Label247"
                                            Caption ="فاصله از پائين كاغذ در چاپ سند :"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =1435
                                    Top =9432
                                    Width =2729
                                    Height =1857
                                    Name ="DAFT_ASN"
                                    SourceObject ="Form.FDAFT_ASN"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2007
                                            Top =9194
                                            Width =1425
                                            Height =240
                                            Name ="DAFT_ASN Label"
                                            Caption ="دفتر اسناد دريافتني"
                                            EventProcPrefix ="DAFT_ASN_Label"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4675
                                    Top =10905
                                    Width =621
                                    Height =450
                                    Name ="Text281"
                                    ControlSource ="CPI"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =5370
                                            Top =10904
                                            Width =2310
                                            Height =435
                                            Name ="Label282"
                                            Caption ="تعداد تاريخهاي قابل برگشت ومشاهده در فاكتور فروش:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6552
                                    Top =11798
                                    Width =1139
                                    Height =227
                                    Name ="SIGN"
                                    ControlSource ="SIGN"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4705
                                            Top =11774
                                            Width =2700
                                            Height =240
                                            Name ="Label311"
                                            Caption ="گردش كاري و فرمها با امضاء عمل شود"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1435
                                    Top =11369
                                    Height =300
                                    BackColor =13434828
                                    Name ="Combo248"
                                    ControlSource ="HESMBAA"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3205
                                            Top =11369
                                            Width =1350
                                            Height =285
                                            Name ="Label320"
                                            Caption ="ماليات:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =3275
                                    Top =11723
                                    Width =1139
                                    Height =227
                                    Name ="CTL_DT"
                                    ControlSource ="CTL_DT"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1428
                                            Top =11699
                                            Width =2700
                                            Height =240
                                            Name ="Label322"
                                            Caption ="تاريخ فقط مربوط به همين سال باشد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =7448
                                    Top =12097
                                    Name ="SAGHF"
                                    ControlSource ="SAGHF"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =3425
                                            Top =12044
                                            Width =3915
                                            Height =255
                                            Name ="Label294"
                                            Caption ="سقف اعتبار اشخاص بر اساس مانده  حساب  كنترل گردد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =7448
                                    Top =12322
                                    Name ="SAGHF2"
                                    ControlSource ="SAGHF2"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =3410
                                            Top =12269
                                            Width =3930
                                            Height =255
                                            Name ="Label315"
                                            Caption ="سقف اعتبار اشخاص بر اساس مانده   اسناد   كنترل گردد"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4672
                                    Top =8870
                                    Width =561
                                    Height =255
                                    Name ="ARSESH"
                                    ControlSource ="ARSESH"
                                    DefaultValue ="0"
                                    InputMask ="#\\%"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =1
                                            Left =5277
                                            Top =8869
                                            Width =1980
                                            Height =240
                                            Name ="Label317"
                                            Caption ="درصد ماليات بر ارزش افزوده:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =1199
                            Top =7560
                            Width =6893
                            Height =5670
                            Name ="Page140"
                            Caption ="فاكتور و انبار"
                            Begin
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =1278
                                    Top =9851
                                    Width =2054
                                    Height =943
                                    Name ="Frame118"
                                    ControlSource ="KALA"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =1471
                                            Top =9693
                                            Width =1710
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label249"
                                            Caption =" كالاها با موجودي صفر"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2844
                                            Top =10089
                                            OptionValue =1
                                            Name ="Option120"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1412
                                                    Top =10041
                                                    Width =1365
                                                    Height =285
                                                    Name ="Label250"
                                                    Caption ="نمايش داده شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =2844
                                            Top =10419
                                            OptionValue =2
                                            Name ="Option122"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1352
                                                    Top =10381
                                                    Width =1425
                                                    Height =285
                                                    Name ="Label251"
                                                    Caption ="نمايش داده نشود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =6258
                                    Top =7781
                                    Width =1755
                                    Height =943
                                    Name ="Frame130"
                                    ControlSource ="WAR"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =6369
                                            Top =7619
                                            Width =1425
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label252"
                                            Caption ="در زمان تغيير مقادير"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =7606
                                            Top =8019
                                            OptionValue =1
                                            Name ="Option132"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =6303
                                                    Top =7962
                                                    Width =1305
                                                    Height =285
                                                    Name ="Label133"
                                                    Caption ="اخطار داده شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =7606
                                            Top =8349
                                            OptionValue =2
                                            Name ="Option134"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =6258
                                                    Top =8359
                                                    Width =1350
                                                    Height =285
                                                    Name ="Label253"
                                                    Caption ="اخطار داده نشود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1426
                                    Top =10834
                                    Width =621
                                    Height =300
                                    Name ="DIG"
                                    ControlSource ="DIG"
                                    ValidationRule ="<14"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2146
                                            Top =10832
                                            Width =2115
                                            Height =285
                                            Name ="Label136"
                                            Caption ="تعداد اعشار در موجودي انبار:"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =3483
                                    Top =9851
                                    Width =2417
                                    Height =943
                                    Name ="Frame141"
                                    ControlSource ="LST"
                                    DefaultValue ="2"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =3625
                                            Top =9693
                                            Width =2025
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label142"
                                            Caption ="در زمان ورود كالاي جديد"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5454
                                            Top =10089
                                            OptionValue =1
                                            Name ="Option143"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =3545
                                                    Top =10045
                                                    Width =1950
                                                    Height =285
                                                    Name ="Label144"
                                                    Caption ="ليست بطوراتوماتيك باز شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5454
                                            Top =10419
                                            OptionValue =2
                                            Name ="Option145"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4370
                                                    Top =10385
                                                    Width =1125
                                                    Height =285
                                                    Name ="Label146"
                                                    Caption ="ليست باز نشود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =7832
                                    Top =9122
                                    Name ="MAND"
                                    ControlSource ="MAND"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =4196
                                            Top =9071
                                            Width =3525
                                            Height =255
                                            Name ="Label183"
                                            Caption ="مانده حساب  اشخاص روي فاكتور نمايش داده شود"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =7267
                                    Top =10908
                                    Name ="MOJU"
                                    ControlSource ="MOJU"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =4321
                                            Top =10847
                                            Width =2820
                                            Height =255
                                            Name ="Label254"
                                            Caption ="موجودي انبار در زمان فروش كنترل شود"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    DecimalPlaces =0
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =4714
                                    Top =11536
                                    Width =2031
                                    Height =313
                                    Name ="DEFANB"
                                    ControlSource ="DEFANB"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                                    ColumnWidths ="0"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6829
                                            Top =11534
                                            Width =1185
                                            Height =240
                                            Name ="Label255"
                                            Caption ="پيش فرض انبار :"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    DecimalPlaces =0
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =1203
                                    Top =11537
                                    Width =1536
                                    Height =313
                                    Name ="DEFTKH"
                                    ControlSource ="DEFTKH"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                                    ColumnWidths ="0"
                                    DefaultValue ="3"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =2827
                                            Top =11540
                                            Width =1755
                                            Height =240
                                            Name ="Label256"
                                            Caption ="پيش فرض نوع مشتري :"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =4188
                                    Top =7781
                                    Width =1725
                                    Height =943
                                    Name ="FRUP"
                                    ControlSource ="FRUP"
                                    DefaultValue ="-1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =4473
                                            Top =7619
                                            Width =1140
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label261"
                                            Caption ="ستون مبلغ "
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5534
                                            Top =8019
                                            OptionValue =0
                                            Name ="Option262"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4306
                                                    Top =7962
                                                    Width =1140
                                                    Height =285
                                                    Name ="Label263"
                                                    Caption ="قابل تغيير باشد"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =5534
                                            Top =8349
                                            Name ="Option264"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4248
                                                    Top =8354
                                                    Width =1200
                                                    Height =285
                                                    Name ="Label265"
                                                    Caption ="قابل تغيير نباشد"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =6243
                                    Top =9855
                                    Width =1715
                                    Height =943
                                    Name ="UPDDATE"
                                    ControlSource ="UPDDATE"
                                    DefaultValue ="-1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =6338
                                            Top =9689
                                            Width =1455
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label267"
                                            Caption ="تاريخ فاكتور"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =7589
                                            Top =10093
                                            Name ="Option268"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =6406
                                                    Top =10036
                                                    Width =1080
                                                    Height =285
                                                    Name ="Label269"
                                                    Caption ="قابل تغيير باشد"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =7589
                                            Top =10423
                                            OptionValue =0
                                            Name ="Option270"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =6248
                                                    Top =10439
                                                    Width =1245
                                                    Height =285
                                                    Name ="Label271"
                                                    Caption ="قابل تغيير نباشد"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6868
                                    Top =8780
                                    Width =1139
                                    Height =227
                                    Name ="CODEVIEW"
                                    ControlSource ="CODEVIEW"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4541
                                            Top =8756
                                            Width =3180
                                            Height =240
                                            Name ="Label308"
                                            Caption ="كد كالاها در زمان چاپ نمايش  داده شوند"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6846
                                    Top =11989
                                    Width =1139
                                    Height =227
                                    Name ="TRANSF"
                                    ControlSource ="TRANSF"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4053
                                            Top =11967
                                            Width =3645
                                            Height =240
                                            Name ="Label326"
                                            Caption ="سابقه تغييرات فاكتورها  بعد از قفل شدن ذخيره گردد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6846
                                    Top =12303
                                    Width =1139
                                    Name ="LOCKFAP"
                                    ControlSource ="LOCKFAP"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =3978
                                            Top =12252
                                            Width =3705
                                            Height =255
                                            Name ="Label328"
                                            Caption ="فاكتور بعداز چاپ قفل گردد"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =1199
                                    Top =7767
                                    Width =2882
                                    Height =1784
                                    Name ="Frame105"
                                    ControlSource ="GHAYM"
                                    DefaultValue ="1"
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =2107
                                            Top =7622
                                            Width =1695
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label106"
                                            Caption ="نمايش قيمت پيش فرض"
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3701
                                            Top =8010
                                            OptionValue =1
                                            Name ="Option108"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =2084
                                                    Top =7967
                                                    Width =1530
                                                    Height =285
                                                    Name ="Label109"
                                                    Caption ="آخرين قيمت فروش"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3701
                                            Top =8340
                                            OptionValue =2
                                            Name ="Option110"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =2369
                                                    Top =8307
                                                    Width =1245
                                                    Height =285
                                                    Name ="Label111"
                                                    Caption ="قيمت تعريف اوليه"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3479
                                            Top =8640
                                            Width =500
                                            OptionValue =4
                                            Name ="Option317"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1259
                                                    Top =8612
                                                    Width =2355
                                                    Height =285
                                                    Name ="Label318"
                                                    Caption ="آخرين قيمت فروش به اين مشتري"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3469
                                            Top =8954
                                            Width =500
                                            OptionValue =5
                                            Name ="Option320"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =1199
                                                    Top =8921
                                                    Width =2400
                                                    Height =285
                                                    Name ="Label321"
                                                    Caption ="قيمت مصوب"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =3716
                                            Top =9288
                                            OptionValue =3
                                            Name ="Option112"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =2984
                                                    Top =9266
                                                    Width =645
                                                    Height =285
                                                    Name ="Label113"
                                                    Caption ="هيچكدام"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =4204
                                    Top =9381
                                    Width =1251
                                    Height =255
                                    Name ="TFCODE_E"
                                    ControlSource ="TFCODE_E"
                                    DefaultValue ="0"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =5484
                                            Top =9380
                                            Width =2250
                                            Height =240
                                            Name ="Label347"
                                            Caption ="برچسب خدمات در چاپ فاكتور :"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =7269
                                    Top =11205
                                    Name ="RMOG"
                                    ControlSource ="RMOG"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =3
                                            Left =4323
                                            Top =11144
                                            Width =2820
                                            Height =255
                                            Name ="Label350"
                                            Caption ="موجودي بصورت قطعي كنترل شود"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =1200
                            Top =7559
                            Width =6885
                            Height =5671
                            Name ="Page159"
                            Caption ="حقوق و دستمزد"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6100
                                    Top =8129
                                    Name ="HEZA"
                                    ControlSource ="HEZA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =6330
                                            Top =8099
                                            Width =675
                                            Height =285
                                            Name ="Label160"
                                            Caption ="اضافه كار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6100
                                    Top =8469
                                    Name ="HPAD"
                                    ControlSource ="HPAD"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =6330
                                            Top =8439
                                            Width =450
                                            Height =285
                                            Name ="Label161"
                                            Caption ="پاداش"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6100
                                    Top =8809
                                    Name ="HOLA"
                                    ControlSource ="HOLA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =6330
                                            Top =8779
                                            Width =630
                                            Height =285
                                            Name ="Label162"
                                            Caption ="حق اولاد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6100
                                    Top =9131
                                    Name ="HKHA"
                                    ControlSource ="HKHA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =6337
                                            Top =9104
                                            Width =1170
                                            Height =285
                                            Name ="Label163"
                                            Caption ="خواربار و مسكن"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6100
                                    Top =9449
                                    Width =261
                                    Name ="HNAH"
                                    ControlSource ="HNAH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =6330
                                            Top =9419
                                            Width =615
                                            Height =285
                                            Name ="Label181"
                                            Caption ="حق نهار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6112
                                    Top =9767
                                    Name ="HJAZ"
                                    ControlSource ="HJAZ"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =6342
                                            Top =9737
                                            Width =645
                                            Height =285
                                            Name ="Label164"
                                            Caption ="حق جذب"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =5860
                                    Top =7852
                                    Width =1772
                                    Height =3578
                                    Name ="Box176"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =5965
                                    Top =7739
                                    Width =1545
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label177"
                                    Caption ="مشمول بيمه و ماليات"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6112
                                    Top =10078
                                    Name ="HRAN"
                                    ControlSource ="HRAN"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6342
                                            Top =10048
                                            Width =555
                                            Height =285
                                            Name ="Label165"
                                            Caption ="راندمان"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6112
                                    Top =10390
                                    Name ="HCON"
                                    ControlSource ="HCON"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6349
                                            Top =10367
                                            Width =975
                                            Height =285
                                            Name ="Label167"
                                            Caption ="شرايط محيط"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6130
                                    Top =10709
                                    Name ="HSAY"
                                    ControlSource ="HSAY"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6360
                                            Top =10679
                                            Width =360
                                            Height =285
                                            Name ="Label166"
                                            Caption ="ساير"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =6115
                                    Top =11054
                                    Name ="HSHI"
                                    ControlSource ="HSHI"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6340
                                            Top =11024
                                            Width =840
                                            Height =285
                                            Name ="Label258"
                                            Caption ="حق شيفت"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =8009
                                    Height =300
                                    BackColor =13434828
                                    Name ="HAZEDAR"
                                    ControlSource ="HAZEDAR"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =3530
                                            Top =8009
                                            Width =1425
                                            Height =285
                                            Name ="Label168"
                                            Caption ="هزينه حقوق اداري:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =8354
                                    Height =300
                                    BackColor =13434828
                                    Name ="HAZTOLID"
                                    ControlSource ="HAZTOLID"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =3530
                                            Top =8354
                                            Width =1425
                                            Height =285
                                            Name ="Label248"
                                            Caption ="هزينه حقوق توليد:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =255
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =8699
                                    Height =300
                                    BackColor =13434828
                                    Name ="HAZFROOSH"
                                    ControlSource ="HAZFROOSH"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =8699
                                            Width =1455
                                            Height =285
                                            Name ="Label257"
                                            Caption ="هزينه حقوق فروش:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =1435
                                    Top =7709
                                    Width =3969
                                    Height =4213
                                    Name ="Box178"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    Left =2400
                                    Top =7559
                                    Width =1875
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label179"
                                    Caption ="سر فصل حسابهاي حقوق"
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =9054
                                    Height =300
                                    BackColor =13434828
                                    Name ="HAZKHADAMAT"
                                    ControlSource ="HAZKHADAMAT"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =9059
                                            Width =1500
                                            Height =285
                                            Name ="Label259"
                                            Caption ="هزينه حقوق خدمات:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =9404
                                    Height =300
                                    BackColor =13434828
                                    Name ="EDABIM"
                                    ControlSource ="EDABIM"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =9438
                                            Width =855
                                            Height =285
                                            Name ="Label169"
                                            Caption ="اداره  بيمه:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =9753
                                    Height =300
                                    BackColor =13434828
                                    Name ="HAZBIM"
                                    ControlSource ="HAZBIM"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =9749
                                            Width =1260
                                            Height =285
                                            Name ="Label170"
                                            Caption ="هزينه 23% بيمه:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =10094
                                    Height =300
                                    BackColor =13434828
                                    Name ="BESHO"
                                    ControlSource ="BESHO"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3534
                                            Top =10094
                                            Width =1320
                                            Height =285
                                            Name ="Label171"
                                            Caption ="بستانكاران حقوق:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =10439
                                    Height =300
                                    BackColor =13434828
                                    Name ="BEDMOS"
                                    ControlSource ="BEDMOS"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =10439
                                            Width =1815
                                            Height =285
                                            Name ="Label172"
                                            Caption ="بدهكاران مساعده وغيره:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =10779
                                    Height =300
                                    BackColor =13434828
                                    Name ="PARDAKH"
                                    ControlSource ="PARDAKH"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3529
                                            Top =10779
                                            Width =1020
                                            Height =285
                                            Name ="Label173"
                                            Caption ="محل پرداخت:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =11129
                                    Height =300
                                    BackColor =13434828
                                    Name ="HAZMALI"
                                    ControlSource ="HAZMALI"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =11129
                                            Width =1710
                                            Height =285
                                            Name ="Label174"
                                            Caption ="اداره دارائي بابت ماليات:"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    ColumnCount =2
                                    ListWidth =4536
                                    Left =1765
                                    Top =11474
                                    Height =300
                                    BackColor =13434828
                                    Name ="PSANDHES"
                                    ControlSource ="PSANDHES"
                                    RowSourceType ="Table/View/StoredProc"
                                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME FROM TDETA_HES"
                                    ColumnWidths ="1701;3402"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3535
                                            Top =11474
                                            Width =1710
                                            Height =285
                                            Name ="Label276"
                                            Caption ="صندوق  پرسنل:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =1510
                                    Top =12003
                                    Width =1671
                                    Height =300
                                    Name ="SAGHFH"
                                    ControlSource ="SAGHFH"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =3244
                                            Top =11999
                                            Width =2160
                                            Height =285
                                            Name ="Label260"
                                            Caption ="حد اقل حقوق مشمول ماليات:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5515
                                    Top =11568
                                    Width =876
                                    Height =300
                                    Name ="SANAVP"
                                    ControlSource ="SANAVP"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6454
                                            Top =11564
                                            Width =1185
                                            Height =240
                                            Name ="Label278"
                                            Caption ="حق سنوات پايه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    ReadingOrder =1
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    Left =5515
                                    Top =11952
                                    Width =876
                                    Height =300
                                    Name ="BON"
                                    ControlSource ="BON"
                                    Format ="#,###"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            Left =6454
                                            Top =11948
                                            Width =1185
                                            Height =240
                                            Name ="Label285"
                                            Caption ="بن:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =1315
                                    Top =11984
                                    Width =186
                                    Height =351
                                    Name ="Command309"
                                    Caption ="Command309"
                                    PictureData = Begin
                                        0x2800000020000000200000000100040000000000000200000000000000000000 ,
                                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x80800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                                        0xffff0000ffffff00777777777777777777777777777777777777777777777777 ,
                                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                                        0x7777777777777777777777777777778777777777777777777777777777777888 ,
                                        0x0777777777777777777777777777888000777777777777777777777777788800 ,
                                        0x0707777777777777777777777788800000707777777777777777777778880000 ,
                                        0x0007077777777777777777778880000000007077777777777777777888000000 ,
                                        0x0000070777777777777777888000000000000070777777777777788800000000 ,
                                        0x0000000707777777777788800000000000000000707777777777770000000000 ,
                                        0x0000000000077777777777777880000000007077777777777777777778800000 ,
                                        0x0000707777777777777777777880000000007077777777777777777778800000 ,
                                        0x0000707777777777777777777880000000007077777777777777777778800000 ,
                                        0x0000707777777777777777777880000000007077777777777777777778800000 ,
                                        0x0000707777777777777777777880000000007077777777777777777778800000 ,
                                        0x0000707777777777777777777770000000000077777777777777777777777777 ,
                                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                                        0x7777777777777777
                                    End
                                    ObjectPalette = Begin
                                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                                        0xc0c0c00080808000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                                        0xffffff0000000000
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =5995
                                    Top =12299
                                    Width =396
                                    Height =189
                                    Name ="Check315"
                                    ControlSource ="TENDAR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =6446
                                            Top =12269
                                            Width =915
                                            Height =240
                                            Name ="Label316"
                                            Caption ="ده درصدي"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =1338
                                    Top =12359
                                    Width =1266
                                    Height =336
                                    Name ="Command314"
                                    Caption ="بيشتر"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =1200
                            Top =7559
                            Width =6885
                            Height =5671
                            Name ="ساير"
                            Begin
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6199
                                    Top =8054
                                    Name ="SA_HOGH"
                                    ControlSource ="SA_HOGH"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =6429
                                            Top =8024
                                            Width =495
                                            Height =240
                                            Name ="Label233"
                                            Caption ="حقوق"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6199
                                    Top =8394
                                    Name ="SA_40EZ"
                                    ControlSource ="SA_40EZ"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =6429
                                            Top =8364
                                            Width =960
                                            Height =240
                                            Name ="Label234"
                                            Caption ="40%اضافه كار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6199
                                    Top =8734
                                    Name ="SA_EZAF"
                                    ControlSource ="SA_EZAF"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =6429
                                            Top =8704
                                            Width =690
                                            Height =240
                                            Name ="Label235"
                                            Caption ="اضافه كار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6199
                                    Top =9074
                                    Name ="SA_PADA"
                                    ControlSource ="SA_PADA"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =6429
                                            Top =9044
                                            Width =480
                                            Height =240
                                            Name ="Label236"
                                            Caption ="پاداش"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6199
                                    Top =9415
                                    Name ="SA_HOLA"
                                    ControlSource ="SA_HOLA"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =6429
                                            Top =9385
                                            Width =675
                                            Height =240
                                            Name ="Label237"
                                            Caption ="حق اولاد"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =6199
                                    Top =9755
                                    Name ="SA_KHAR"
                                    ControlSource ="SA_KHAR"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =6429
                                            Top =9725
                                            Width =1155
                                            Height =240
                                            Name ="Label238"
                                            Caption ="خواربار و مسكن"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =4444
                                    Top =8039
                                    Name ="SA_NAHA"
                                    ControlSource ="SA_NAHA"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =4674
                                            Top =8009
                                            Width =615
                                            Height =240
                                            Name ="Label239"
                                            Caption ="حق نهار"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =4435
                                    Top =8381
                                    Name ="SA_JAZB"
                                    ControlSource ="SA_JAZB"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =4665
                                            Top =8351
                                            Width =705
                                            Height =240
                                            Name ="Label240"
                                            Caption ="حق جذب"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =4435
                                    Top =8721
                                    Name ="SA_RAND"
                                    ControlSource ="SA_RAND"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =4665
                                            Top =8691
                                            Width =555
                                            Height =240
                                            Name ="Label241"
                                            Caption ="راندمان"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =4435
                                    Top =9061
                                    Name ="SA_COND"
                                    ControlSource ="SA_COND"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =4665
                                            Top =9031
                                            Width =945
                                            Height =240
                                            Name ="Label242"
                                            Caption ="شرايط محيط"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =4435
                                    Top =9401
                                    Name ="SA_SAYE"
                                    ControlSource ="SA_SAYE"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =4665
                                            Top =9371
                                            Width =405
                                            Height =240
                                            Name ="Label262"
                                            Caption ="ساير"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    ReadingOrder =1
                                    Left =4435
                                    Top =9742
                                    Name ="SA_23BI"
                                    ControlSource ="SA_23BI"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =255
                                            Left =4665
                                            Top =9712
                                            Width =645
                                            Height =240
                                            Name ="Label264"
                                            Caption ="23%بيمه"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    Left =4210
                                    Top =7672
                                    Width =3515
                                    Height =2561
                                    Name ="Box245"
                                End
                                Begin Label
                                    Visible = NotDefault
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =4270
                                    Top =7559
                                    Width =3285
                                    Height =285
                                    BackColor =-2147483633
                                    Name ="Label266"
                                    Caption ="عواملي كه در نرخ يك ساعت كار بايد حذف شوند"
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =1465
                                    Top =7765
                                    Name ="ECONM"
                                    ControlSource ="ECONM"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            Left =1700
                                            Top =7679
                                            Width =2370
                                            Height =300
                                            Name ="Label268"
                                            Caption ="مواد اوليه در زمان توليد خارج  شود"
                                        End
                                    End
                                End
                                Begin OptionGroup
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =4317
                                    Top =10456
                                    Width =3571
                                    Height =1078
                                    Name ="FINALS"
                                    ControlSource ="FINALS"
                                    DefaultValue ="-1"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextAlign =2
                                            Left =4839
                                            Top =10287
                                            Width =2505
                                            Height =285
                                            BackColor =-2147483633
                                            Name ="Label273"
                                            Caption ="خروج مواد به كالاي درجريان "
                                        End
                                        Begin OptionButton
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =7447
                                            Top =10734
                                            Width =245
                                            Name ="Option274"
                                            Begin
                                                Begin Label
                                                    Visible = NotDefault
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =5679
                                                    Top =10677
                                                    Width =1650
                                                    Height =285
                                                    Name ="Label275"
                                                    Caption ="بصورت كلي انجام شود"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            ReadingOrder =1
                                            Left =7447
                                            Top =11064
                                            Width =335
                                            OptionValue =0
                                            Name ="Option276"
                                            Begin
                                                Begin Label
                                                    Visible = NotDefault
                                                    OverlapFlags =247
                                                    TextFontCharSet =0
                                                    TextAlign =3
                                                    Left =4512
                                                    Top =11074
                                                    Width =2820
                                                    Height =285
                                                    Name ="Label277"
                                                    Caption =" به معين كالاي ساخته شده منتقل شود"
                                                End
                                            End
                                        End
                                    End
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    ReadingOrder =1
                                    Left =1368
                                    Top =8035
                                    Width =275
                                    Name ="SANAT"
                                    ControlSource ="SANAT"
                                    Begin
                                        Begin Label
                                            Visible = NotDefault
                                            OverlapFlags =247
                                            Left =1708
                                            Top =8024
                                            Width =2370
                                            Height =285
                                            Name ="Label303"
                                            Caption ="ارتيكل اسناد صنعتي صادر شود"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =8681
                                    Width =1086
                                    Name ="STFR"
                                    ControlSource ="STFR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2958
                                            Top =8684
                                            Width =930
                                            Height =240
                                            Name ="Label323"
                                            Caption ="فاكتور فروش"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =9021
                                    Width =1086
                                    Name ="STKH"
                                    ControlSource ="STKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2953
                                            Top =9024
                                            Width =885
                                            Height =240
                                            Name ="Label324"
                                            Caption ="فاكتور خريد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =9361
                                    Width =1086
                                    Name ="STHFR"
                                    ControlSource ="STHFR"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2953
                                            Top =9354
                                            Width =960
                                            Height =240
                                            Name ="Label325"
                                            Caption ="حواله فروش:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =9702
                                    Width =1086
                                    Name ="STHKH"
                                    ControlSource ="STHKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2950
                                            Top =9702
                                            Width =900
                                            Height =240
                                            Name ="Label327"
                                            Caption ="رسيد انبار:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =10042
                                    Width =1086
                                    Name ="STENT"
                                    ControlSource ="STENT"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2950
                                            Top =10042
                                            Width =885
                                            Height =240
                                            Name ="Label329"
                                            Caption ="انتقالي:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =10382
                                    Width =1086
                                    Name ="STKHS"
                                    ControlSource ="STKHS"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2950
                                            Top =10382
                                            Width =885
                                            Height =240
                                            Name ="Label330"
                                            Caption ="خروج ساير:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =10722
                                    Width =1086
                                    Name ="STKHH"
                                    ControlSource ="STKHH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2950
                                            Top =10722
                                            Width =900
                                            Height =240
                                            Name ="Label331"
                                            Caption ="خروج:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =11062
                                    Width =1086
                                    Name ="STTOL"
                                    ControlSource ="STTOL"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2950
                                            Top =11062
                                            Width =885
                                            Height =240
                                            Name ="Label332"
                                            Caption ="توليد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =11402
                                    Width =1086
                                    Name ="STFRB"
                                    ControlSource ="STFRB"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2953
                                            Top =11409
                                            Width =1125
                                            Height =240
                                            Name ="Label333"
                                            Caption ="برگشت فروش:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1531
                                    Top =11743
                                    Width =1086
                                    Name ="STBKH"
                                    ControlSource ="STBKH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =2953
                                            Top =11739
                                            Width =1020
                                            Height =240
                                            Name ="Label334"
                                            Caption ="برگشت خريد:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =1263
                                    Top =8519
                                    Width =2826
                                    Height =4145
                                    Name ="Box341"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =2
                                    Left =2163
                                    Top =8369
                                    Width =1035
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Label342"
                                    Caption ="شماره شروع "
                                    ControlTipText ="قبل از صدور سند در هر بخش ميتوانيد شماره شروع سند را در ذيل مشخص كنيد"
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1518
                                    Top =12044
                                    Width =1086
                                    Name ="STMO"
                                    ControlSource ="STMO"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2943
                                            Top =12044
                                            Width =1155
                                            Height =240
                                            Name ="Label344"
                                            Caption ="رسيد مستقيم:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    ReadingOrder =2
                                    ScrollBarAlign =2
                                    NumeralShapes =2
                                    IMESentenceMode =3
                                    Left =1518
                                    Top =12333
                                    Width =1086
                                    Name ="STKHA"
                                    ControlSource ="STKHA"
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2943
                                            Top =12329
                                            Width =1035
                                            Height =240
                                            Name ="Label346"
                                            Caption ="فاكتور خدمات:"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =1200
                            Top =7560
                            Width =6885
                            Height =5670
                            Name ="Page283"
                            Caption ="ISO"
                            Begin
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =4774
                                    Top =8139
                                    Width =1191
                                    Height =330
                                    Name ="ISO_FROOSH"
                                    ControlSource ="ISO_FROOSH"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =6025
                                            Top =8132
                                            Width =1260
                                            Height =330
                                            Name ="Label287"
                                            Caption ="حواله انبار فروش:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =4774
                                    Top =8556
                                    Width =1191
                                    Height =330
                                    Name ="ISO_KHAREED"
                                    ControlSource ="ISO_KHAREED"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =6025
                                            Top =8549
                                            Width =1260
                                            Height =330
                                            Name ="Label289"
                                            Caption ="رسيد انبار خريد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =4774
                                    Top =8973
                                    Width =1191
                                    Height =330
                                    Name ="ISO_MAVAD"
                                    ControlSource ="ISO_MAVAD"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =6025
                                            Top =8969
                                            Width =1590
                                            Height =345
                                            Name ="Label291"
                                            Caption ="حواله خروج مواد اوليه:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =4774
                                    Top =9387
                                    Width =1191
                                    Height =330
                                    Name ="ISO_TOLID"
                                    ControlSource ="ISO_TOLID"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =6025
                                            Top =9383
                                            Width =1590
                                            Height =345
                                            Name ="Label293"
                                            Caption ="رسيد انبار  آمار توليد:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =255
                                    IMESentenceMode =3
                                    Left =4774
                                    Top =9801
                                    Width =1191
                                    Height =330
                                    Name ="ISO_MAVADSAYER"
                                    ControlSource ="ISO_MAVADSAYER"
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =6025
                                            Top =9797
                                            Width =1590
                                            Height =345
                                            Name ="Label295"
                                            Caption ="حواله خروج ساير مواد:"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =4328
                                    Top =7906
                                    Width =3401
                                    Height =2607
                                    Name ="Box296"
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    Left =6491
                                    Top =7792
                                    Width =1125
                                    Height =240
                                    BackColor =-2147483633
                                    Name ="Label297"
                                    Caption ="كدينگ فرمهاي "
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_AUTOMATIC.cls"
