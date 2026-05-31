Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15397
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =375
    Top =2415
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1302cc5f0caee240
    End
    InputParameters ="@Forms___F_MENU_HOTEL___DT1=[Forms]![F_MENU_HOTEL]![DT1]"
    RecordSource ="R_LIST_MEHMAN"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,112 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x48004f00540045004c005f005f005f004400540031003d005b0046006f007200 ,
        0x6d0073005d0021005b0046005f004d0045004e0055005f0048004f0054004500 ,
        0x4c005d0021005b004400540031005d00
    End
    RecordSourceQualifier ="dbo"
    Begin
        Begin Line
            Width =1701
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =13797
                    Width =1095
                    Height =300
                    FontSize =10
                    Name ="MNAME"
                    ControlSource ="MNAME"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =11539
                    Width =772
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =9049
                    Width =666
                    Height =300
                    FontSize =10
                    TabIndex =2
                    Name ="KNOW"
                    ControlSource ="KNOW"
                    FontName ="Zar"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =14940
                    Width =456
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =10319
                    Width =666
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="BDATE"
                    ControlSource ="BDATE"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =9769
                    Width =516
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="TABEI"
                    ControlSource ="TABEI"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =8375
                    Width =621
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="MSODUR"
                    ControlSource ="MSODUR"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =7505
                    Width =831
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="JOB"
                    ControlSource ="JOB"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =6574
                    Width =876
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="GHASD"
                    ControlSource ="GHASD"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =5734
                    Width =771
                    Height =300
                    FontSize =10
                    TabIndex =9
                    Name ="SCITY"
                    ControlSource ="SCITY"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =4894
                    Width =786
                    Height =300
                    FontSize =10
                    TabIndex =10
                    Name ="DCITY"
                    ControlSource ="DCITY"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =4024
                    Width =786
                    Height =300
                    FontSize =10
                    TabIndex =11
                    Name ="ENGIN"
                    ControlSource ="ENGIN"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =3049
                    Width =936
                    Height =300
                    FontSize =10
                    TabIndex =12
                    Name ="VDATE"
                    ControlSource ="VDATEE"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1849
                    Width =546
                    Height =300
                    FontSize =10
                    TabIndex =13
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    FontName ="Zar"
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =11295
                    Top =105
                    Width =215
                    Height =195
                    TabIndex =14
                    Name ="Check159"
                    ControlSource ="=Not [SEX]"
                End
                Begin CheckBox
                    OldBorderStyle =0
                    Left =11040
                    Top =105
                    Width =155
                    Height =195
                    TabIndex =15
                    Name ="Check161"
                    ControlSource ="SEX"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1189
                    Width =591
                    Height =300
                    FontSize =10
                    TabIndex =16
                    Name ="ROOMN"
                    ControlSource ="ROOMN"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =2449
                    Width =561
                    Height =300
                    FontSize =10
                    TabIndex =17
                    Name ="HAMRAH"
                    ControlSource ="HAMRAH"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =12344
                    Width =1417
                    Height =300
                    FontSize =10
                    TabIndex =18
                    Name ="MFAMILY"
                    ControlSource ="MFAMILY"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    BackStyle =0
                    Left =34
                    Width =1086
                    Height =300
                    FontSize =10
                    TabIndex =19
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Zar"
                End
            End
        End
    End
End
CodeBehindForm
' See "HOTEL_SUB1.cls"
