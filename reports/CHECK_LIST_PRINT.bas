Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =102
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10161
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =600
    Top =930
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x5d881081fadee340
    End
    RecordSource ="SELECT PAY_GETD.N_SERI, PAY_GETD.N_HESAB, PAY_GETD.DATE_S, TCOD_BANKS.NAMES + N'"
        " ' + PAY_GETD.SHOBEH AS BANSHO, PAY_GETD.MABL, PAY_GETD.CUST_NO, TDETA_HES.NAME,"
        " CHRE_LSPH.RADIF AS list_no, PAY_GETD.RADIF, PAY_GETD.BANK FROM PAY_GETD INNER J"
        "OIN TCOD_BANKS ON PAY_GETD.BANK = TCOD_BANKS.CODE INNER JOIN TDETA_HES ON PAY_GE"
        "TD.N_KOL = TDETA_HES.N_KOL AND PAY_GETD.N_MOIN = TDETA_HES.NUMBER AND PAY_GETD.N"
        "_TAF = TDETA_HES.TNUMBER INNER JOIN CHRE_LSPH ON PAY_GETD.N_SERI = CHRE_LSPH.N_S"
        "ERI AND PAY_GETD.BANK = CHRE_LSPH.BANK AND PAY_GETD.DATE_S = CHRE_LSPH.DATE_S"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin BreakLevel
            ControlSource ="BANK"
        End
        Begin BreakLevel
            ControlSource ="N_SERI"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =4762
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8970
                    Top =2130
                    Width =1134
                    Height =465
                    FontSize =10
                    FontWeight =700
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7035
                    Top =2130
                    Width =1911
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5580
                    Top =2130
                    Width =1418
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="Koodak Mazar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2325
                    Top =2130
                    Width =3156
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="BANSHO"
                    ControlSource ="BANSHO"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Top =2137
                    Width =2181
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =1
                    IMESentenceMode =3
                    Top =105
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="Text5"
                    ControlSource ="=FARSIDATE(Date())"
                    FontName ="Koodak Mazar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =2775
                    Width =8166
                    Height =465
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="Text6"
                    ControlSource ="=ALPHANUM([MABL])"
                    Format ="#,###"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7590
                    Top =3285
                    Width =1374
                    Height =465
                    FontSize =10
                    TabIndex =7
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4170
                    Top =3285
                    Width =1479
                    Height =465
                    FontSize =10
                    TabIndex =8
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2655
                    Top =3750
                    Width =1374
                    Height =465
                    FontSize =10
                    TabIndex =9
                    Name ="list_no"
                    ControlSource ="list_no"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Top =570
                    Width =4551
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Koodak Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8775
                    Top =3765
                    Width =1374
                    Height =465
                    FontSize =10
                    TabIndex =11
                    Name ="Text11"
                    ControlSource ="=[Page]"
                    FontName ="Koodak Mazar"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHECK_LIST_PRINT.cls"
