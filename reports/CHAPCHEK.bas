Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =10780
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =600
    Top =930
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc77cdd20b9efe240
    End
    Caption ="چك قابل چاپ"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextFontFamily =2
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
            Height =6927
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    BackStyle =0
                    Left =5852
                    Width =2336
                    Height =495
                    FontSize =14
                    FontWeight =700
                    Name ="DATE"
                    ControlSource ="=Forms!CHAPCHEK!DATE_S"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =5250
                    Top =1455
                    Width =916
                    Height =780
                    FontSize =13
                    FontWeight =700
                    TabIndex =1
                    Name ="TEXTMABL"
                    ControlSource ="=\"*\" & ALPHANUM(Forms!CHAPCHEK!MABL) & \" ريال********************************"
                        "**********************************************\""
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =5713
                    Top =2594
                    Width =686
                    Height =1620
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="VAJH"
                    ControlSource ="=Forms!CHAPCHEK!VAJH & \"*******************************************************"
                        "************\""
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =2400
                    Top =540
                    Width =2433
                    Height =510
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="MABL"
                    ControlSource ="=Forms!CHAPCHEK!MABL"
                    Format ="\"-/\"#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =15
                    BackStyle =0
                    Left =7031
                    Top =615
                    Width =210
                    Height =660
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    Name ="DATEN"
                    ControlSource ="=ALPHANUM(Right([Forms]![CHAPCHEK]![DATE_S],2)) & \" / \" & ALPHANUM(Mid([Forms]"
                        "![CHAPCHEK]![DATE_S],5,2)) & \" / \" & ALPHANUM(Left([Forms]![CHAPCHEK]![DATE_S]"
                        ",4))"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHAPCHEK.cls"
