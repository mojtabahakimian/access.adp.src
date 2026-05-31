Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3973
    DatasheetFontHeight =10
    ItemSuffix =58
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x08f4fe4eecb5e240
    End
    RecordSource ="SELECT_HOKM"
    Caption ="SELECT HOKM"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =9991
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3045
                            Width =810
                            Height =240
                            Name ="Label34"
                            Caption ="كد پرسنل:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =340
                    TabIndex =1
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3405
                            Top =340
                            Width =450
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =680
                    TabIndex =2
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2760
                            Top =680
                            Width =1095
                            Height =240
                            Name ="Label36"
                            Caption ="روزهاي كاركرد:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =1021
                    TabIndex =3
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3105
                            Top =1021
                            Width =750
                            Height =240
                            Name ="Label37"
                            Caption ="اضافه كار:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =1361
                    TabIndex =4
                    Name ="PADASH"
                    ControlSource ="PADASH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3315
                            Top =1361
                            Width =540
                            Height =240
                            Name ="Label38"
                            Caption ="پاداش:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =1701
                    TabIndex =5
                    Name ="KASR_VAM"
                    ControlSource ="KASR_VAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =1701
                            Width =735
                            Height =240
                            Name ="Label39"
                            Caption ="كسر وام:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1894
                    Top =2041
                    TabIndex =6
                    Name ="SEX"
                    ControlSource ="SEX"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1294
                            Top =2011
                            Width =630
                            Height =240
                            Name ="Label40"
                            Caption ="جنسيت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1894
                    Top =2381
                    TabIndex =7
                    Name ="TAB56"
                    ControlSource ="TAB56"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1219
                            Top =2351
                            Width =705
                            Height =240
                            Name ="Label41"
                            Caption ="تبصره اي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =2721
                    TabIndex =8
                    Name ="HDATE"
                    ControlSource ="HDATE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3045
                            Top =2721
                            Width =810
                            Height =240
                            Name ="Label42"
                            Caption ="تاريخ حكم:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =3061
                    TabIndex =9
                    Name ="SALARY_DAYLY"
                    ControlSource ="SALARY_DAYLY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2880
                            Top =3061
                            Width =975
                            Height =240
                            Name ="Label43"
                            Caption ="حقوق روزانه:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =3402
                    TabIndex =10
                    Name ="SALARY_DAYLYB"
                    ControlSource ="SALARY_DAYLYB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2880
                            Top =3402
                            Width =975
                            Height =240
                            Name ="Label44"
                            Caption ="حقوق روزانه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =3742
                    TabIndex =11
                    Name ="HOME"
                    ControlSource ="HOME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2340
                            Top =3742
                            Width =1515
                            Height =240
                            Name ="Label45"
                            Caption ="حق خواربار و مسكن:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =4082
                    TabIndex =12
                    Name ="CHILDREN"
                    ControlSource ="CHILDREN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =4082
                            Width =735
                            Height =240
                            Name ="Label46"
                            Caption ="حق اولاد:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =4422
                    TabIndex =13
                    Name ="CONDITIONS"
                    ControlSource ="CONDITIONS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2625
                            Top =4422
                            Width =1230
                            Height =240
                            Name ="Label47"
                            Caption ="شرايط محيط كار:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =4762
                    TabIndex =14
                    Name ="BON"
                    ControlSource ="BON"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3555
                            Top =4762
                            Width =300
                            Height =240
                            Name ="Label48"
                            Caption ="بن:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =5102
                    TabIndex =15
                    Name ="JAZB"
                    ControlSource ="JAZB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3090
                            Top =5102
                            Width =765
                            Height =240
                            Name ="Label49"
                            Caption ="حق جذب:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =5443
                    TabIndex =16
                    Name ="SAYER"
                    ControlSource ="SAYER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3390
                            Top =5443
                            Width =465
                            Height =240
                            Name ="Label50"
                            Caption ="ساير:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =5783
                    TabIndex =17
                    Name ="HAZCENTER"
                    ControlSource ="HAZCENTER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2970
                            Top =5783
                            Width =885
                            Height =240
                            Name ="Label51"
                            Caption ="مركز هزينه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =6123
                    TabIndex =18
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2625
                            Top =6123
                            Width =1230
                            Height =240
                            Name ="Label52"
                            Caption ="نام  مركز مصرف:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =6463
                    TabIndex =19
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3240
                            Top =6463
                            Width =615
                            Height =240
                            Name ="Label53"
                            Caption ="كد  كالا:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =6803
                    TabIndex =20
                    Name ="MABEZ"
                    ControlSource ="MABEZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2760
                            Top =6803
                            Width =1095
                            Height =240
                            Name ="Label54"
                            Caption ="مبلغ اضافه كار:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =453
                    Top =7143
                    TabIndex =21
                    Name ="MABEZT"
                    ControlSource ="MABEZT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2760
                            Top =7143
                            Width =1095
                            Height =240
                            Name ="Label55"
                            Caption ="مبلغ اضافه كار:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =7484
                    TabIndex =22
                    Name ="EZAD"
                    ControlSource ="EZAD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2685
                            Top =7484
                            Width =1170
                            Height =240
                            Name ="Label56"
                            Caption ="اضافه كار عادي:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =453
                    Top =7824
                    TabIndex =23
                    Name ="EZTAT"
                    ControlSource ="EZTAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2910
                            Top =7824
                            Width =945
                            Height =240
                            Name ="Label57"
                            Caption ="تعطيل كاري:"
                        End
                    End
                End
            End
        End
    End
End
