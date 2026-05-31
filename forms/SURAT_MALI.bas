Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =119
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15363
    DatasheetFontHeight =10
    ItemSuffix =113
    Left =3720
    Top =390
    Right =19080
    Bottom =8580
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x5463b96e0e97e440
    End
    RecordSource ="SURAT_MALI_HEAD"
    Caption ="تعريف و برپاسازي صورت مالي"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8243
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    SpecialEffect =1
                    BackStyle =1
                    OverlapFlags =85
                    Left =6497
                    Top =446
                    Width =3827
                    Height =7692
                    BackColor =16764057
                    Name ="Frame4"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7961
                            Top =165
                            Width =1065
                            Height =240
                            BackColor =-2147483633
                            Name ="Label5"
                            Caption ="عناوين ترازنامه"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6602
                            Top =523
                            OptionValue =0
                            Name ="Option100"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6922
                                    Top =495
                                    Width =2235
                                    Height =240
                                    Name ="Label101"
                                    Caption ="شماره سند افتتاحيه"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6688
                            Top =808
                            OptionValue =1
                            Name ="Option7"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6918
                                    Top =780
                                    Width =2325
                                    Height =240
                                    Name ="Label8"
                                    Caption ="موجودي مواد اوليه در ابتداي دوره"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =1112
                            OptionValue =2
                            Name ="Option9"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =1084
                                    Width =1725
                                    Height =240
                                    Name ="Label10"
                                    Caption ="خريد مواد اوليه طي دوره"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =1397
                            OptionValue =3
                            Name ="Option11"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =1369
                                    Width =1275
                                    Height =240
                                    Name ="Label12"
                                    Caption ="دستمزد مستقيم"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =1682
                            OptionValue =4
                            Name ="Option13"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =1654
                                    Width =2430
                                    Height =240
                                    Name ="Label14"
                                    Caption ="ساير هزينه هاي توليد(سربار توليد)"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =1967
                            OptionValue =5
                            Name ="Option15"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =1939
                                    Width =2775
                                    Height =240
                                    Name ="Label16"
                                    Caption ="موجودي كالاي در جريان ساخت اول دوره"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =2252
                            OptionValue =6
                            Name ="Option17"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =2224
                                    Width =1905
                                    Height =240
                                    Name ="Label18"
                                    Caption ="موجودي كالا در ابتداي دوره"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =2522
                            OptionValue =7
                            Name ="Option19"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =2494
                                    Width =1305
                                    Height =240
                                    Name ="Label20"
                                    Caption ="خريد كالا طي دوره"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =2807
                            OptionValue =8
                            Name ="Option21"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =2779
                                    Width =405
                                    Height =240
                                    Name ="Label22"
                                    Caption ="ساير"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =3092
                            OptionValue =9
                            Name ="Option23"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =3064
                                    Width =1365
                                    Height =240
                                    Name ="Label24"
                                    Caption ="موجودي نزد ديگران"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =3377
                            OptionValue =10
                            Name ="Option25"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =3349
                                    Width =1500
                                    Height =240
                                    Name ="Label26"
                                    Caption ="موجودي كالاي اماني"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =3662
                            OptionValue =11
                            Name ="Option27"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =3634
                                    Width =1050
                                    Height =240
                                    Name ="Label28"
                                    Caption ="ساير موجوديها"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =3947
                            OptionValue =12
                            Name ="Option29"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =3919
                                    Width =1020
                                    Height =240
                                    Name ="Label30"
                                    Caption ="فروش ناخالص"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =4232
                            OptionValue =13
                            Name ="Option31"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =4204
                                    Width =1950
                                    Height =240
                                    Name ="Label32"
                                    Caption ="برگشت از فروش و تخفيفات"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =4517
                            OptionValue =14
                            Name ="Option33"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =4489
                                    Width =1545
                                    Height =240
                                    Name ="Label34"
                                    Caption ="درآمد حاصل از خدمات"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =4802
                            OptionValue =15
                            Name ="Option35"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =4774
                                    Width =3225
                                    Height =240
                                    Name ="Label36"
                                    Caption ="درآمد حاصل از معاملات بانكي و سرمايه گذاري"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =5087
                            OptionValue =16
                            Name ="Option37"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =5059
                                    Width =1410
                                    Height =240
                                    Name ="Label38"
                                    Caption ="درآمد حاصل از اجاره"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =5357
                            OptionValue =17
                            Name ="Option39"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =5329
                                    Width =1575
                                    Height =240
                                    Name ="Label40"
                                    Caption ="در آمدهاي غير مترقبه"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =5642
                            OptionValue =18
                            Name ="Option41"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =5614
                                    Width =1875
                                    Height =240
                                    Name ="Label42"
                                    Caption ="ساير دريافتي ها و درآمدها"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =5927
                            OptionValue =19
                            Name ="Option43"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =5899
                                    Width =1215
                                    Height =240
                                    Name ="Label44"
                                    Caption ="حقوق و دستمزد"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6673
                            Top =6212
                            OptionValue =20
                            Name ="Option45"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =6903
                                    Top =6184
                                    Width =1245
                                    Height =240
                                    Name ="Label46"
                                    Caption ="هزينه هاي اداري"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6587
                            Top =6497
                            OptionValue =21
                            Name ="Option47"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6903
                                    Top =6469
                                    Width =2325
                                    Height =240
                                    Name ="Label92"
                                    Caption ="هزينه هاي مالي"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6587
                            Top =6782
                            OptionValue =22
                            Name ="Option93"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6903
                                    Top =6754
                                    Width =2325
                                    Height =240
                                    Name ="Label94"
                                    Caption ="هزينه هاي توزيع و فروش"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6587
                            Top =7067
                            OptionValue =23
                            Name ="Option95"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6903
                                    Top =7039
                                    Width =2325
                                    Height =240
                                    Name ="Label96"
                                    Caption ="هزينه هاي استهلاك"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6587
                            Top =7352
                            OptionValue =24
                            Name ="Option97"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6903
                                    Top =7324
                                    Width =2325
                                    Height =240
                                    Name ="Label98"
                                    Caption ="ساير هزينه ها"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =6587
                            Top =7700
                            OptionValue =25
                            Name ="Option102"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6907
                                    Top =7672
                                    Width =2235
                                    Height =240
                                    Name ="Label103"
                                    Caption ="تا شماره سند  محاسبه كن"
                                End
                            End
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =453
                    Width =6180
                    Height =6855
                    TabIndex =1
                    Name ="SURAT_MALI_SUB"
                    SourceObject ="Form.SURAT_MALI_SUB"
                    LinkChildFields ="CODG;prg"
                    LinkMasterFields ="FRAME4;prg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =2443
                            Top =170
                            Width =1395
                            Height =240
                            Name ="SURAT_MALI_SUB Label"
                            Caption ="حسابهاي موثر"
                            EventProcPrefix ="SURAT_MALI_SUB_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1585
                    Top =7420
                    Width =3118
                    Height =454
                    TabIndex =2
                    Name ="Command99"
                    Caption ="مشاهده صورت مالي"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13425
                    Top =450
                    Width =921
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    Name ="prg"
                    ControlSource ="prg"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14607
                            Top =457
                            Width =585
                            Height =330
                            Name ="Label106"
                            Caption ="شماره:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10545
                    Top =915
                    Width =3801
                    Height =345
                    TabIndex =4
                    BackColor =13434828
                    Name ="ONVAN"
                    ControlSource ="ONVAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14607
                            Top =922
                            Width =585
                            Height =330
                            Name ="Label108"
                            Caption ="عنوان:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12645
                    Top =1380
                    Height =345
                    TabIndex =5
                    Name ="DATE_E"
                    ControlSource ="DATE_E"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14607
                            Top =1387
                            Width =585
                            Height =330
                            Name ="Label110"
                            Caption ="تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12645
                    Top =1845
                    Height =345
                    TabIndex =6
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14607
                            Top =1852
                            Width =585
                            Height =330
                            Name ="Label112"
                            Caption ="كاربر:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SURAT_MALI.cls"
