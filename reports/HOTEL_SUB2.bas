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
    Width =6859
    DatasheetFontHeight =10
    Left =375
    Top =3045
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_HOTEL___DT1=[Forms]![F_MENU_HOTEL]![DT1]"
    RecSrcDt = Begin
        0xcce24e304dafe240
    End
    InputParameters ="@Forms___F_MENU_HOTEL___DT1=[Forms]![F_MENU_HOTEL]![DT1]"
    RecordSource ="R_LIST_MEHMAN_KHORUG"
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
                    Left =5247
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
                    Left =2989
                    Width =772
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="TABEI1"
                    ControlSource ="TABEI"
                    Format ="Standard"
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
                    Left =6390
                    Width =456
                    Height =300
                    FontSize =10
                    TabIndex =2
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
                    Left =2025
                    Width =936
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="VDATE"
                    ControlSource ="VDATE"
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
                    Left =825
                    Width =546
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Width =756
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="MODAT"
                    ControlSource ="MODAT"
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
                    Left =1425
                    Width =561
                    Height =300
                    FontSize =10
                    TabIndex =6
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
                    Left =3794
                    Width =1417
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="MFAMILY"
                    ControlSource ="MFAMILY"
                    Format ="Standard"
                    FontName ="Zar"
                End
            End
        End
    End
End
CodeBehindForm
' See "HOTEL_SUB2.cls"
