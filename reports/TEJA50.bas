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
    Width =10771
    DatasheetFontHeight =10
    ItemSuffix =11
    Left =705
    Top =1035
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6efb6d2932cce240
    End
    RecordSource ="SELECT DATE_S, MABL, NAME_TAH, N_TAF AS N_MOIN, N_SERI FROM CHK_SODUR GROUP BY D"
        "ATE_S, MABL, NAME_TAH, N_TAF, N_SERI"
    Caption ="چك قابل چاپ"
    OnOpen ="[Event Procedure]"
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
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    BackStyle =0
                    NumeralShapes =2
                    Left =8175
                    Width =1811
                    Height =495
                    FontSize =14
                    FontWeight =700
                    Name ="DATE"
                    ControlSource ="DATE_S"
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
                    Top =1065
                    Width =2896
                    Height =915
                    FontSize =13
                    FontWeight =700
                    TabIndex =1
                    Name ="TEXTMABL"
                    ControlSource ="=\"*\" & ALPHANUM([MABL]) & \" ريال*********************************************"
                        "*********************************\""
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =5775
                    Top =2060
                    Width =3746
                    Height =975
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="VAJH"
                    ControlSource ="=[NAME_TAH] & \"****************************************************************"
                        "***\""
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
                    Width =2328
                    Height =495
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="\"-/\"#,###"
                    FontName ="Zar"
                End
            End
        End
    End
End
CodeBehindForm
' See "TEJA50.cls"
