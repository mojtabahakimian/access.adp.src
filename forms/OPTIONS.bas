Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14685
    DatasheetFontHeight =10
    ItemSuffix =249
    Left =285
    Top =2910
    Right =14970
    Bottom =10665
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xdb9b41e868efe340
    End
    Caption ="تنظيمات"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =9210
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CheckBox
                    OverlapFlags =93
                    Left =14030
                    Top =222
                    Name ="Check9"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7204
                            Top =165
                            Width =6720
                            Height =240
                            Name ="Label25"
                            Caption ="اولين انبار معرفي شده در بخش كليد انبارها به عنوان انبار پيش فرض براي هر كاربر ن"
                                "شان داده شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =14027
                    Top =612
                    TabIndex =1
                    Name ="Check10"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11056
                            Top =555
                            Width =2865
                            Height =240
                            Name ="Label27"
                            Caption ="كد مخصوص چاپ فاكتور خاص اين شركت:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =14027
                    Top =1002
                    TabIndex =2
                    Name ="Check13"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8626
                            Top =945
                            Width =5295
                            Height =540
                            Name ="Label33"
                            Caption ="در زمان تعريف سرفصلهاي فروش آن را بصورت فروش: - 1-فروش داخلي 2-فروش صادراتي 3- ت"
                                "خفيفات فروش 4-برگشت فروش تعريف كن"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =14003
                    Top =1589
                    TabIndex =3
                    Name ="Check14"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11079
                            Top =1530
                            Width =2820
                            Height =240
                            Name ="Label35"
                            Caption ="حساب ارزي فعال شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =14005
                    Top =1947
                    TabIndex =4
                    Name ="Check15"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8535
                            Top =1890
                            Width =5370
                            Height =240
                            Name ="Label37"
                            Caption ="فيلد توضيحات و کد فني در تعريف كالا در زمان جستجوي كالا  نمايش داده نشود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =14005
                    Top =2262
                    TabIndex =5
                    Name ="Check16"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8874
                            Top =2205
                            Width =5025
                            Height =240
                            Name ="Label39"
                            Caption ="صدور فاكتور فقط براي حواله هايي كه تيك بار گيري خورده امكان پذير باشد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =14011
                    Top =2592
                    TabIndex =6
                    Name ="Check17"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8640
                            Top =2535
                            Width =5265
                            Height =240
                            Name ="Label41"
                            Caption ="صدور رسيد خريد كالا فقط بر اساس درخواست كالا امكان پذير باشد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10201
                    Top =570
                    Width =807
                    Height =255
                    TabIndex =7
                    Name ="C1112"
                    ValidationRule ="<99 And >-1"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =93
                    AccessKey =199
                    TextFontCharSet =0
                    Left =10080
                    Top =6975
                    Width =1290
                    Height =345
                    FontWeight =700
                    TabIndex =8
                    Name ="Command147"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =208
                    TextFontCharSet =0
                    Left =8805
                    Top =6975
                    Width =1275
                    Height =345
                    FontWeight =700
                    TabIndex =9
                    Name ="Command148"
                    Caption ="&ذخيره و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =14017
                    Top =3702
                    TabIndex =10
                    Name ="Check23"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13936
                    Top =3297
                    TabIndex =11
                    Name ="Check19"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8085
                            Top =3240
                            Width =5835
                            Height =240
                            Name ="Label57"
                            Caption ="در فاكتور فروش نوع مشتري بر اساس تعريف نوع مشتري بصورت خودكار مقدار بگيرد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14275
                    Top =113
                    Width =340
                    Height =283
                    FontWeight =700
                    Name ="Label54"
                    Caption ="1"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14275
                    Top =540
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label58"
                    Caption ="2"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14275
                    Top =960
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label56"
                    Caption ="3"
                End
                Begin Label
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =1545
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label59"
                    Caption ="4"
                End
                Begin Label
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =1905
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label60"
                    Caption ="5"
                End
                Begin Label
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =2220
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label61"
                    Caption ="6"
                End
                Begin Label
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =2550
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label62"
                    Caption ="7"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =3240
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label63"
                    Caption ="9"
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =3660
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label64"
                    Caption ="10"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =2880
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label65"
                    Caption ="8"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =4050
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label66"
                    Caption ="11"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =4710
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label67"
                    Caption ="13"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =5040
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label69"
                    Caption ="14"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9276
                    Top =3645
                    Width =4647
                    Height =330
                    TabIndex =12
                    Name ="C2425"
                    RowSourceType ="Value List"
                    RowSource ="00;\"كارت ساعت جهانگستر\";01;\"كارت ساعتPW\";02;\"كارت ساعت آراز\""
                    ColumnWidths ="0"
                    ValidationRule ="<99 And >-1"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13936
                    Top =2937
                    TabIndex =13
                    Name ="Check18"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11520
                            Top =2880
                            Width =2385
                            Height =240
                            Name ="Label78"
                            Caption ="فاكتور فروش مستقيم عمل شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13942
                    Top =4122
                    TabIndex =14
                    Name ="Check26"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =11931
                            Top =4065
                            Width =2010
                            Height =240
                            Name ="Label80"
                            Caption ="اندازه فونت در فاكتور فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11031
                    Top =4050
                    Width =807
                    Height =255
                    TabIndex =15
                    Name ="C2728"
                    ValidationRule ="<99 And >0"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13969
                    Top =5070
                    Width =235
                    TabIndex =16
                    Name ="Check30"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10302
                            Top =5052
                            Width =3615
                            Height =240
                            Name ="Label85"
                            Caption ="كليه حسابها در فيلد  محل مصرف  نشان داده شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =13926
                    Top =4466
                    Width =284
                    Height =284
                    TabIndex =17
                    Name ="Check29"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =9516
                            Top =4410
                            Width =4410
                            Height =240
                            Name ="Label87"
                            Caption ="فاكتور فروش بدون انبار و با كمترين  تعداد اينتر باشد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =4410
                    Width =340
                    Height =240
                    FontWeight =700
                    Name ="Label88"
                    Caption ="12"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =14022
                    Top =4769
                    Width =179
                    Height =284
                    TabIndex =18
                    Name ="Check7"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =10345
                            Top =4712
                            Width =3570
                            Height =240
                            Name ="Label90"
                            Caption ="پنجره جستجوي كالا در فاكتور  با اينتر خالي باز نشود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13951
                    Top =5442
                    TabIndex =19
                    Name ="Check38"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9015
                            Top =5385
                            Width =4905
                            Height =240
                            Name ="Label106"
                            Caption ="حساب صندوق (نقد)درخزانه داري بر اساس واحد و شيفت  انتخاب شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =5355
                    Width =340
                    Height =285
                    FontWeight =700
                    Name ="Label112"
                    Caption ="15"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =13924
                    Top =5733
                    Width =284
                    Height =284
                    TabIndex =20
                    Name ="Check39"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =10952
                            Top =5670
                            Width =2985
                            Height =240
                            Name ="Label114"
                            Caption ="مدت زمان باز شدن پنجره ورود كاربر(Login)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10077
                    Top =5681
                    Width =912
                    Height =255
                    TabIndex =21
                    Name ="C4041"
                    ValidationRule =">0"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14244
                    Top =5715
                    Width =340
                    Height =240
                    FontWeight =700
                    Name ="Label116"
                    Caption ="16"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =9525
                    Top =5674
                    Width =495
                    Height =285
                    FontWeight =700
                    Name ="Label141"
                    Caption ="دقيقه"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =13920
                    Top =6014
                    Width =284
                    Height =284
                    TabIndex =22
                    Name ="Check42"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =10695
                            Top =5955
                            Width =3225
                            Height =240
                            Name ="Label144"
                            Caption ="مانده حساب ها در خزانه داري نشان داده شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14235
                    Top =5955
                    Width =345
                    Height =285
                    FontWeight =700
                    Name ="Label145"
                    Caption ="17"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =13935
                    Top =6299
                    Width =284
                    Height =284
                    TabIndex =23
                    Name ="Check43"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =10710
                            Top =6240
                            Width =3225
                            Height =240
                            Name ="Label149"
                            Caption ="مركز هزينه فعال شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14235
                    Top =6240
                    Width =375
                    Height =285
                    FontWeight =700
                    Name ="Label150"
                    Caption ="18"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13935
                    Top =6599
                    Width =284
                    Height =284
                    TabIndex =24
                    Name ="Check44"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =9780
                            Top =6540
                            Width =4155
                            Height =240
                            Name ="Label152"
                            Caption ="پنجره سند حسابداري بر اساس رزولوشن مانيتور تنظيم شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =14220
                    Top =6540
                    Width =405
                    Height =285
                    FontWeight =700
                    Name ="Label153"
                    Caption ="19"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =177
                    Width =261
                    TabIndex =25
                    Name ="Check32"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =120
                            Width =5355
                            Height =240
                            Name ="Label89"
                            Caption ="در فاكتور فروش نوع مشتري در سند حسابداري  ذكر شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =552
                    Width =261
                    TabIndex =26
                    Name ="Check33"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =495
                            Width =5355
                            Height =240
                            Name ="Label91"
                            Caption ="در فاكتور فروش قيمت مصرف كننده نشان داده شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =120
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label92"
                    Caption ="22"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =495
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label93"
                    Caption ="23"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =1245
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label96"
                    Caption ="25"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =1283
                    Width =261
                    Height =284
                    TabIndex =27
                    Name ="Check34"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =135
                            Top =1275
                            Width =5355
                            Height =240
                            Name ="Label100"
                            Caption ="انتخاب نوع ترازوي ديجيتال"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =405
                    Top =1260
                    Width =3237
                    Height =330
                    TabIndex =28
                    Name ="C3536"
                    RowSourceType ="Value List"
                    RowSource ="00;\"ويستا\";01;\"صدرا\";02;\"سي اس2\";03;\"پند\""
                    ColumnWidths ="0"
                    StatusBarText ="00"
                    ValidationRule ="<99 And >-1"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =840
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label107"
                    Caption ="24"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =908
                    Width =261
                    Height =284
                    TabIndex =29
                    Name ="Check37"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =900
                            Width =5355
                            Height =240
                            Name ="Label109"
                            Caption ="قيمت مصرف كننده از اف دو بيايد وگرنه از آخرين فاكتور ثبت شده  ميگيرد"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =601
                    Top =8725
                    TabIndex =30
                    Name ="Check40"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            TextAlign =3
                            Left =396
                            Top =8668
                            Width =261
                            Height =240
                            Name ="Label111"
                            Caption ="."
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =1650
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label156"
                    Caption ="26"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =2025
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label160"
                    Caption ="27"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =14186
                    Top =7030
                    Width =261
                    TabIndex =31
                    Name ="Check161"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =9240
                            Top =6973
                            Width =4920
                            Height =240
                            Name ="Label118"
                            Caption =" \"تخفيف بر اساس نحوه پرداخت\" در تب ساير در فاكتور فروش چاپ شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =2355
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label120"
                    Caption ="28"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =14186
                    Top =7750
                    Width =261
                    TabIndex =32
                    Name ="Check162"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9240
                            Top =7693
                            Width =4920
                            Height =240
                            Name ="Label122"
                            Caption ="پيش فرض واحد وشيفت را  از آخرين فاكتور نگيرد"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =14186
                    Top =7405
                    Width =261
                    TabIndex =33
                    Name ="Check163"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9240
                            Top =7348
                            Width =4920
                            Height =240
                            Name ="Label124"
                            Caption ="ستون كد كالا در فاكتور فروش نشان داده شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =2730
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label125"
                    Caption ="29"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =3090
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label126"
                    Caption ="30"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =1707
                    Width =261
                    TabIndex =34
                    Name ="Check45"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =1650
                            Width =5355
                            Height =240
                            Name ="Label128"
                            Caption ="تبديل پيش فاكتور به فاكتور  اجازه دهد كه  تكراري فاكتور صادر كند"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =3435
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label129"
                    Caption ="31"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =1992
                    Width =261
                    TabIndex =35
                    Name ="Check46"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =1935
                            Width =5355
                            Height =405
                            Name ="Label131"
                            Caption ="در زمان تبديل پيش فاكتور به فاكتور نام كاربر صادر كننده پيش فاكتور در ملاحظات فا"
                                "كتور ذكر شود."
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =3705
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label132"
                    Caption ="32"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =2787
                    Width =261
                    TabIndex =36
                    Name ="Check48"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =135
                            Top =2730
                            Width =5355
                            Height =240
                            Name ="Label134"
                            Caption ="كيبورد اصلاح شود براي حرف ي و ك"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =2442
                    Width =261
                    TabIndex =37
                    Name ="Check47"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =135
                            Top =2385
                            Width =5355
                            Height =405
                            Name ="Label136"
                            Caption ="ستون تخفيفات نقدي علاوه بر ستون تخفيفات  نمايش داده شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =4038
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label137"
                    Caption ="33"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =4575
                    Width =435
                    Height =240
                    FontWeight =700
                    Name ="Label138"
                    Caption ="35"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =3177
                    Width =261
                    TabIndex =38
                    Name ="Check49"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =3120
                            Width =5355
                            Height =240
                            Name ="Label140"
                            Caption ="فعال شدن دريافت شماره سريال كالا "
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =3462
                    Width =261
                    Height =284
                    TabIndex =39
                    Name ="Check50"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =135
                            Top =3405
                            Width =5355
                            Height =240
                            Name ="Label166"
                            Caption ="مقدار كالا در فاكتور فروش  صفر قبول نكند"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =1282
                    Top =3685
                    TabIndex =40
                    Name ="Check1"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1077
                            Top =3628
                            Width =261
                            Height =240
                            Name ="Label169"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =1282
                    Top =4102
                    TabIndex =41
                    Name ="Check2"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1077
                            Top =4045
                            Width =261
                            Height =240
                            Name ="Label171"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =1282
                    Top =4519
                    TabIndex =42
                    Name ="Check3"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1077
                            Top =4462
                            Width =261
                            Height =240
                            Name ="Label173"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =1282
                    Top =4936
                    TabIndex =43
                    Name ="Check4"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1077
                            Top =4879
                            Width =261
                            Height =240
                            Name ="Label175"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =1282
                    Top =5353
                    TabIndex =44
                    Name ="Check5"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1077
                            Top =5296
                            Width =261
                            Height =240
                            Name ="Label177"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2075
                    Top =3912
                    TabIndex =45
                    Name ="Check6"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1870
                            Top =3855
                            Width =261
                            Height =240
                            Name ="Label179"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2075
                    Top =4329
                    TabIndex =46
                    Name ="Check180"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1870
                            Top =4272
                            Width =261
                            Height =240
                            Name ="Label181"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2075
                    Top =4746
                    TabIndex =47
                    Name ="Check8"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1870
                            Top =4689
                            Width =261
                            Height =240
                            Name ="Label183"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2075
                    Top =5163
                    TabIndex =48
                    Name ="Check184"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1870
                            Top =5106
                            Width =261
                            Height =240
                            Name ="Label185"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2075
                    Top =5580
                    TabIndex =49
                    Name ="Check11"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =1870
                            Top =5523
                            Width =261
                            Height =240
                            Name ="Label187"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2869
                    Top =4252
                    TabIndex =50
                    Name ="Check12"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =2664
                            Top =4195
                            Width =261
                            Height =240
                            Name ="Label189"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2869
                    Top =4669
                    TabIndex =51
                    Name ="Check20"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =2664
                            Top =4612
                            Width =261
                            Height =240
                            Name ="Label191"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2869
                    Top =5086
                    TabIndex =52
                    Name ="Check21"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =2664
                            Top =5029
                            Width =261
                            Height =240
                            Name ="Label193"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =2869
                    Top =5503
                    TabIndex =53
                    Name ="Check22"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =2664
                            Top =5446
                            Width =261
                            Height =240
                            Name ="Label31"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =3890
                    Top =4308
                    TabIndex =54
                    Name ="Check31"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =3685
                            Top =4251
                            Width =261
                            Height =240
                            Name ="Label197"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =3890
                    Top =4725
                    TabIndex =55
                    Name ="Check35"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =3685
                            Top =4668
                            Width =261
                            Height =240
                            Name ="Label199"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =3890
                    Top =5142
                    TabIndex =56
                    Name ="Check36"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextAlign =3
                            Left =3685
                            Top =5085
                            Width =261
                            Height =240
                            Name ="Label201"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =3747
                    Width =261
                    TabIndex =57
                    Name ="Check52"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =3690
                            Width =5355
                            Height =240
                            Name ="Label202"
                            Caption ="جايزه در نرم افزار فعال باشد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =4047
                    Width =261
                    TabIndex =58
                    Name ="Check53"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =3990
                            Width =5355
                            Height =240
                            Name ="Label204"
                            Caption ="فاكتور فروش مستقيم براي هركاربر از تعيين دسترسي انتخاب شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =4320
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label205"
                    Caption ="34"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =4332
                    Width =261
                    TabIndex =59
                    Name ="Check54"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =4275
                            Width =5355
                            Height =240
                            Name ="Label207"
                            Caption ="خلاصه و فشرده  سازي فاكتور فعال شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =4587
                    Width =261
                    TabIndex =60
                    Name ="Check55"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =4530
                            Width =5355
                            Height =240
                            Name ="Label209"
                            Caption ="در فاكتور فروش واحد فروش در سند حسابداري  ذكر شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =4860
                    Width =435
                    Height =240
                    FontWeight =700
                    Name ="Label210"
                    Caption ="36"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =4872
                    Width =261
                    TabIndex =61
                    Name ="Check56"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =4815
                            Width =5355
                            Height =240
                            Name ="Label212"
                            Caption ="چند فرموله عمل كند"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =5175
                    Width =435
                    Height =240
                    FontWeight =700
                    Name ="Label216"
                    Caption ="37"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =5187
                    Width =261
                    TabIndex =62
                    Name ="Check58"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =5130
                            Width =5355
                            Height =240
                            Name ="Label218"
                            Caption ="تخفيفات پيشرفته عمل كند"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =5952
                    Width =261
                    TabIndex =63
                    Name ="Check59"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =5895
                            Width =5355
                            Height =240
                            Name ="Label220"
                            Caption ="در زمان تبديل پيش فاكتور به فاكتور موجودي كنترل گردد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =5895
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label146"
                    Caption ="39"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =5502
                    Width =261
                    TabIndex =64
                    Name ="Check51"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =135
                            Top =5445
                            Width =5355
                            Height =405
                            Name ="Label221"
                            Caption ="در زمان تبديل پيش فاكتور به فاكتور مبلغ فروش كنترل گردد كه اگر تغيير كرده اخطار "
                                "دهد بجز مبالغ صفر"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =5490
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label151"
                    Caption ="38"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5532
                    Top =6239
                    Width =261
                    Height =300
                    TabIndex =65
                    Name ="Check60"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =135
                            Top =6240
                            Width =5355
                            Height =240
                            ForeColor =255
                            Name ="Label223"
                            Caption ="مد محافظت شده فعال شود(نرم افزار کندتر ميشود)"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =6210
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label224"
                    Caption ="40"
                End
                Begin CheckBox
                    OverlapFlags =247
                    ReadingOrder =1
                    Left =5532
                    Top =6525
                    Width =261
                    Height =300
                    TabIndex =66
                    Name ="Check61"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =135
                            Top =6525
                            Width =5355
                            Height =345
                            Name ="Label176"
                            Caption ="موجودي بر اساس کاردکس در جستجوي کالا نشان داده شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =6510
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label225"
                    Caption ="41"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =6855
                    Width =261
                    Height =300
                    TabIndex =67
                    Name ="Check62"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =135
                            Top =6855
                            Width =5355
                            Height =240
                            Name ="Label227"
                            Caption ="ماليات بر ارزش افزوده  جزء پورسانت لحاظ گردد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =6840
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label228"
                    Caption ="42"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =13560
                    Top =8447
                    Width =261
                    Height =300
                    TabIndex =68
                    Name ="Check63"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =8163
                            Top =8447
                            Width =5355
                            Height =240
                            Name ="Label230"
                            Caption ="ماليات بر ارزش افزوده  جزء پورسانت لحاظ گردد"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =12597
                    Top =8050
                    Width =261
                    Height =300
                    TabIndex =69
                    Name ="Check57"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =7200
                            Top =8050
                            Width =5355
                            Height =240
                            Name ="Label232"
                            Caption ="."
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5532
                    Top =7155
                    Width =261
                    Height =300
                    TabIndex =70
                    Name ="Check64"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =135
                            Top =7155
                            Width =5355
                            Height =405
                            Name ="Label235"
                            Caption ="در زمان صدور صورت حساب (f8ياCtrl+f8)كنترل كند كه فاكتورهاي فروش سند خورده باشد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5805
                    Top =7140
                    Width =435
                    Height =285
                    FontWeight =700
                    Name ="Label236"
                    Caption ="43"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5532
                    Top =7545
                    Width =261
                    Height =300
                    TabIndex =71
                    Name ="Check65"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =135
                            Top =7545
                            Width =5355
                            Height =405
                            Name ="Label238"
                            Caption ="SSM  فعال شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5790
                    Top =7530
                    Width =465
                    Height =285
                    FontWeight =700
                    Name ="Label239"
                    Caption ="44"
                End
                Begin CheckBox
                    OverlapFlags =255
                    ReadingOrder =1
                    Left =5532
                    Top =7830
                    Width =261
                    Height =300
                    TabIndex =72
                    Name ="Check66"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =1785
                            Top =7830
                            Width =3705
                            Height =240
                            Name ="Label241"
                            Caption ="انبار کالاي ساخته شده ميانگين موزون استاندارد باشد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =119
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5775
                    Top =7815
                    Width =465
                    Height =285
                    FontWeight =700
                    Name ="Label242"
                    Caption ="45"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5532
                    Top =8145
                    Width =261
                    Height =300
                    TabIndex =73
                    Name ="Check67"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1305
                            Top =8145
                            Width =4185
                            Height =240
                            Name ="Label244"
                            Caption ="تاييد اسناد قبل از ورود سند فاکتور فروش و خروج مواد و توليد"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5760
                    Top =8130
                    Width =495
                    Height =285
                    FontWeight =700
                    Name ="Label245"
                    Caption ="46"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =5532
                    Top =8520
                    Width =261
                    Height =300
                    TabIndex =74
                    Name ="Check68"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1305
                            Top =8520
                            Width =4185
                            Height =240
                            Name ="Label247"
                            Caption ="پارامترهاي بيشتر در خريد نشان داده شود"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =5745
                    Top =8505
                    Width =525
                    Height =285
                    FontWeight =700
                    Name ="Label248"
                    Caption ="47"
                End
            End
        End
    End
End
CodeBehindForm
' See "OPTIONS.cls"
