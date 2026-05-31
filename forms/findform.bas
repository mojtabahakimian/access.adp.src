Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =4305
    Top =7455
    Right =11505
    Bottom =9090
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcbd2eb4509fde140
    End
    Caption ="جستـــــــــجو"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =1647
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =93
                    Top =106
                    Width =1206
                    Height =373
                    FontWeight =700
                    TabIndex =1
                    Name ="Command1"
                    Caption ="&اولين عبارت"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1605
                    Top =165
                    Width =4131
                    Height =345
                    Name ="findwhat"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            ReadingOrder =2
                            Left =5848
                            Top =178
                            Width =1125
                            Height =285
                            Name ="Label3"
                            Caption ="عبارت جستجو:"
                            FontName ="Tahoma"
                            ControlTipText ="كلمه يا جمله ا ي كه مايليد به دنبال آن بگرديد.ميتوانيد از علامت سوال يا *  استفا"
                                "ده نماييد."
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =218
                    TextFontCharSet =0
                    Left =93
                    Top =599
                    Width =1206
                    Height =373
                    FontWeight =700
                    TabIndex =2
                    Name ="Command4"
                    Caption ="&عبارت بعدي"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =204
                    TextFontCharSet =0
                    Left =93
                    Top =1092
                    Width =1206
                    Height =373
                    FontWeight =700
                    TabIndex =3
                    Name ="Command5"
                    Caption ="پايان &جستجو"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    Left =5850
                    Top =645
                    Width =1350
                    Height =285
                    Name ="Label6"
                    Caption =" جستجو به سمت:"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    Left =5865
                    Top =1095
                    Width =1110
                    Height =285
                    Name ="Label7"
                    Caption ="نحوه  انطباق:"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3968
                    Top =627
                    Width =1772
                    Height =255
                    TabIndex =4
                    Name ="searchfrom"
                    RowSourceType ="Value List"
                    RowSource ="1;\"بالا\";2;\"پايين\";3;\"همه جا\""
                    ColumnWidths ="0"
                    DefaultValue ="3"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3968
                    Top =1077
                    Width =1772
                    Height =270
                    TabIndex =5
                    Name ="matchform"
                    RowSourceType ="Value List"
                    RowSource ="1;\"ابتداي فيلد\";2;\"تمام فيلد\";3;\"هرقسمتي از فيلد\""
                    ColumnWidths ="0"
                    DefaultValue ="3"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3753
                    Top =628
                    Width =170
                    Height =210
                    TabIndex =6
                    Name ="matchcase"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1425
                            Top =585
                            Width =2310
                            Height =285
                            Name ="Label13"
                            Caption ="انطبــــاق  بااندازه حروف"
                            FontName ="Tahoma"
                            ControlTipText ="اگر اين بخش را علامت بزنيد حروف كوچك وبزرگ متفاوت درنظر گرفته ميشود."
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3755
                    Top =963
                    Width =170
                    Height =210
                    TabIndex =7
                    Name ="searchall"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1425
                            Top =900
                            Width =2265
                            Height =285
                            Name ="Label15"
                            Caption ="فيلدهايي كه با فرمت هستند"
                            FontName ="Tahoma"
                            ControlTipText ="فيلد هايي كه داراي حالات مشخص هستند مانند عدد  22,1,55  كه داراي فرمت است وعدد 2"
                                "2155 كه بدون فرمت است"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3729
                    Top =1258
                    Width =170
                    Height =210
                    TabIndex =8
                    Name ="onlycurren"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =3
                            Left =1305
                            Top =1170
                            Width =2385
                            Height =285
                            FontSize =7
                            Name ="Label17"
                            Caption ="جستجو فقط درهمين فيلد انجام  شود"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "findform.cls"
