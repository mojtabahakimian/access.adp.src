Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =249
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11451
    DatasheetFontHeight =10
    ItemSuffix =32
    Left =8505
    Top =2400
    Right =19950
    Bottom =8910
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x477a3ed0d2d3e540
    End
    Caption ="جذب انحراف نرخ و مصرف به قيمت تمام شده"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =746
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =3855
                    Top =56
                    Width =3675
                    Height =285
                    FontSize =10
                    Name ="Label0"
                    Caption ="تبديل قيمت استاندارد به تاريخي (واقعلي)"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3855
                    Top =461
                    Width =3675
                    Height =285
                    FontSize =10
                    Name ="Label6"
                    Caption ="Standard cost to date cost "
                End
            End
        End
        Begin Section
            Height =5782
            BackColor =-2147483633
            Name ="Detail"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8844
                    Top =113
                    Width =1755
                    Height =330
                    BackColor =16777164
                    Name ="FMON"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON ORDER BY MON_ID"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10794
                            Top =118
                            Width =495
                            Height =285
                            Name ="Label3"
                            Caption ="ماه"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4478
                    Top =4535
                    Width =1984
                    Height =624
                    TabIndex =1
                    Name ="Command1"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8845
                    Top =935
                    TabIndex =2
                    Name ="copf"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9082
                            Top =907
                            Width =840
                            Height =240
                            Name ="Label4"
                            Caption ="کپي فرمول"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =4260
                    Top =120
                    Width =2841
                    Height =300
                    TabIndex =3
                    Name ="ANBARMAS"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR WHERE (CODE <> 0) ORDER BY NAMES"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7207
                            Top =127
                            Width =1605
                            Height =285
                            Name ="Label5"
                            Caption ="انبار مواد مصرفي توليد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8850
                    Top =1243
                    TabIndex =4
                    Name ="uptolidf"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =9087
                            Top =1215
                            Width =2190
                            Height =240
                            Name ="Label9"
                            Caption ="آبديت توليدات با فرمول جديد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8844
                    Top =1560
                    TabIndex =5
                    Name ="UPGHEYMAT"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9081
                            Top =1532
                            Width =2190
                            Height =240
                            Name ="Label11"
                            Caption ="آبديت قيمت مواد در فرمول"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8845
                    Top =1890
                    TabIndex =6
                    Name ="UPMEGH"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =9082
                            Top =1862
                            Width =2190
                            Height =240
                            Name ="Label13"
                            Caption ="ابديت مقادير فرمول "
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8845
                    Top =2239
                    TabIndex =7
                    Name ="UPKHORUG"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =9084
                            Top =2207
                            Width =2280
                            Height =240
                            Name ="Label131"
                            Caption ="ابديت خروج مواد طبق فرمول "
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8850
                    Top =2522
                    TabIndex =8
                    Name ="BAZSA"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =9089
                            Top =2490
                            Width =2280
                            Height =240
                            Name ="Label19"
                            Caption ="بازسازي نرخ ميانگين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4592
                    Top =2437
                    Width =1755
                    Height =330
                    TabIndex =9
                    BackColor =16777164
                    Name ="TEDADP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6540
                            Top =2445
                            Width =885
                            Height =285
                            Name ="Label22"
                            Caption ="تعدا پردازش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4590
                    Top =2895
                    Width =1755
                    Height =330
                    TabIndex =10
                    BackColor =16777164
                    Name ="CO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6538
                            Top =2903
                            Width =885
                            Height =285
                            Name ="Label25"
                            Caption ="کد کالا"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8850
                    Top =2914
                    TabIndex =11
                    Name ="anb"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =9089
                            Top =2882
                            Width =2280
                            Height =240
                            Name ="Label27"
                            Caption ="آبديت انبار گرداني"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1079
                    Top =1077
                    Width =1755
                    Height =330
                    TabIndex =12
                    BackColor =10092543
                    Name ="dast"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3030
                            Top =1080
                            Width =1380
                            Height =285
                            Name ="Label29"
                            Caption ="دستمزد جذب شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1077
                    Top =1535
                    Width =1755
                    Height =330
                    TabIndex =13
                    BackColor =10092543
                    Name ="sarb"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3030
                            Top =1545
                            Width =1215
                            Height =285
                            Name ="Label31"
                            Caption ="سربار جذب شده"
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
' See "zjazbenheraf.cls"
