Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5167
    RowHeight =375
    ItemSuffix =14
    Left =7905
    Top =1965
    Right =13320
    Bottom =5820
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xba27363991e8e540
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS TAG1, HEAD_LST.N_S, HEAD_LST.USER_NAME, "
        "HEAD_LST.DATE_N, HEAD_LST.CUST_NO, CUST_HESAB.NAME, TAGCOD.BARGAH, COUNT(INVO_LS"
        "T.NUMBER) AS ARTI FROM HEAD_LST INNER JOIN TAGCOD ON HEAD_LST.TAG = TAGCOD.CODE "
        "LEFT OUTER JOIN INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG ="
        " INVO_LST.TAG LEFT OUTER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_HESAB.hes GR"
        "OUP BY HEAD_LST.N_S, HEAD_LST.USER_NAME, HEAD_LST.DATE_N, HEAD_LST.NUMBER, HEAD_"
        "LST.TAG, HEAD_LST.CUST_NO, CUST_HESAB.NAME, TAGCOD.BARGAH HAVING (COUNT(INVO_LST"
        ".NUMBER) = 0) AND (HEAD_LST.TAG <> 3) AND (HEAD_LST.TAG <> 4 AND HEAD_LST.TAG <>"
        " 12 AND HEAD_LST.TAG <> 13 AND HEAD_LST.TAG <> 25 AND HEAD_LST.TAG <> 27)"
    Caption ="برگه هاي خالي"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin Section
            Height =3868
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2154
                    Top =453
                    ColumnWidth =990
                    ColumnOrder =6
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =340
                            Top =453
                            Width =885
                            Height =240
                            Name ="Label3"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =680
                    Top =1757
                    ColumnWidth =6045
                    ColumnOrder =3
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Top =1757
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =680
                    Top =2097
                    ColumnWidth =1110
                    ColumnOrder =4
                    TabIndex =2
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3362
                            Top =2097
                            Width =720
                            Height =240
                            Name ="Label8"
                            Caption ="كد "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =680
                    Top =2437
                    ColumnWidth =1605
                    ColumnOrder =5
                    TabIndex =3
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3047
                            Top =2437
                            Width =1035
                            Height =240
                            Name ="Label9"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =690
                    Top =1410
                    ColumnWidth =1080
                    ColumnOrder =2
                    TabIndex =4
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3219
                            Top =1410
                            Width =810
                            Height =240
                            Name ="Label10"
                            Caption ="*تاريخ "
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =906
                    Top =2607
                    ColumnWidth =1815
                    ColumnOrder =1
                    TabIndex =5
                    Name ="BARGAH"
                    ControlSource ="BARGAH"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3618
                            Top =2607
                            Width =690
                            Height =240
                            Name ="Label11"
                            Caption ="نوع برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =736
                    Top =737
                    ColumnWidth =960
                    ColumnOrder =0
                    TabIndex =6
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3118
                            Top =737
                            Width =1020
                            Height =240
                            Name ="Label12"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =623
                    Top =3628
                    TabIndex =7
                    Name ="TAG1"
                    ControlSource ="TAG1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3335
                            Top =3628
                            Width =690
                            Height =240
                            Name ="Label13"
                            Caption ="برچسب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FREE_FACTOR.cls"
