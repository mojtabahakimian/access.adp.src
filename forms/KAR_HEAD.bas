Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =1
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =32
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =30
    GridY =30
    Width =20144
    DatasheetFontHeight =10
    ItemSuffix =57
    Top =540
    Right =15120
    Bottom =7155
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xce29db6ad2bde240
    End
    RecordSource ="KAR_HEAD"
    Caption ="كاربرگ ثبت كاركرد پرسنل"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName =""
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            Width =1701
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8314
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =17283
                    Top =56
                    Width =1086
                    Height =284
                    BackColor =13434828
                    Name ="KNUM"
                    ControlSource ="KNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18425
                            Top =60
                            Width =1080
                            Height =284
                            Name ="Label1"
                            Caption ="شماره كاربرگ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11060
                    Top =61
                    Width =3105
                    Height =284
                    TabIndex =2
                    BackColor =13434828
                    Name ="MODIR"
                    ControlSource ="MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14233
                            Top =60
                            Width =795
                            Height =284
                            Name ="Label5"
                            Caption ="مدير توليد:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9005
                    Top =56
                    Width =1470
                    Height =284
                    TabIndex =3
                    BackColor =13434828
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10536
                            Top =56
                            Width =435
                            Height =284
                            Name ="Label7"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11060
                    Top =435
                    Width =7305
                    Height =284
                    TabIndex =4
                    BackColor =13434828
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =18483
                            Top =435
                            Width =825
                            Height =284
                            Name ="Label9"
                            Caption ="توضيحات:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =15095
                    Top =60
                    Height =284
                    TabIndex =1
                    BackColor =13434828
                    Name ="KMON"
                    ControlSource ="KMON"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON ORDER BY MON_ID"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =16846
                            Top =64
                            Width =345
                            Height =284
                            Name ="Label3"
                            Caption ="ماه:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =45
                    Top =915
                    Width =20040
                    Height =6960
                    TabIndex =5
                    Name ="KAR_BARG_SUB"
                    SourceObject ="Form.KAR_BARG_SUB"
                    LinkChildFields ="KNUM;KMON"
                    LinkMasterFields ="KNUM;KMON"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =16234
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =6
                    BackColor =13434828
                    Name ="Text17"
                    ControlSource ="=[KAR_BARG_SUB]![SD1]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =17109
                            Top =7928
                            Width =1095
                            Height =285
                            Name ="Label18"
                            Caption ="جمع كاركردماه:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =15754
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =7
                    BackColor =13434828
                    Name ="Text20"
                    ControlSource ="=[KAR_BARG_SUB]![SD2]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =15274
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =8
                    BackColor =13434828
                    Name ="Text22"
                    ControlSource ="=[KAR_BARG_SUB]![SD3]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =14794
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =9
                    BackColor =13434828
                    Name ="Text23"
                    ControlSource ="=[KAR_BARG_SUB]![SD4]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =14314
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =10
                    BackColor =13434828
                    Name ="Text24"
                    ControlSource ="=[KAR_BARG_SUB]![SD5]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13834
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =11
                    BackColor =13434828
                    Name ="Text25"
                    ControlSource ="=[KAR_BARG_SUB]![SD6]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13354
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =12
                    BackColor =13434828
                    Name ="Text26"
                    ControlSource ="=[KAR_BARG_SUB]![SD7]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12874
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =13
                    BackColor =13434828
                    Name ="Text27"
                    ControlSource ="=[KAR_BARG_SUB]![SD8]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12394
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =14
                    BackColor =13434828
                    Name ="Text28"
                    ControlSource ="=[KAR_BARG_SUB]![SD9]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11914
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =15
                    BackColor =13434828
                    Name ="Text29"
                    ControlSource ="=[KAR_BARG_SUB]![SD10]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11434
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =16
                    BackColor =13434828
                    Name ="Text30"
                    ControlSource ="=[KAR_BARG_SUB]![SD11]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10954
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =17
                    BackColor =13434828
                    Name ="Text31"
                    ControlSource ="=[KAR_BARG_SUB]![SD12]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10474
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =18
                    BackColor =13434828
                    Name ="Text32"
                    ControlSource ="=[KAR_BARG_SUB]![SD13]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9994
                    Top =7921
                    Width =449
                    Height =284
                    TabIndex =19
                    BackColor =13434828
                    Name ="Text33"
                    ControlSource ="=[KAR_BARG_SUB]![SD14]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9499
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =20
                    BackColor =13434828
                    Name ="Text35"
                    ControlSource ="=[KAR_BARG_SUB]![SD15]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9004
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =21
                    BackColor =13434828
                    Name ="Text36"
                    ControlSource ="=[KAR_BARG_SUB]![SD16]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8494
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =22
                    BackColor =13434828
                    Name ="Text37"
                    ControlSource ="=[KAR_BARG_SUB]![SD17]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7999
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =23
                    BackColor =13434828
                    Name ="Text38"
                    ControlSource ="=[KAR_BARG_SUB]![SD18]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7504
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =24
                    BackColor =13434828
                    Name ="Text39"
                    ControlSource ="=[KAR_BARG_SUB]![SD19]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7009
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =25
                    BackColor =13434828
                    Name ="Text40"
                    ControlSource ="=[KAR_BARG_SUB]![SD20]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6514
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =26
                    BackColor =13434828
                    Name ="Text41"
                    ControlSource ="=[KAR_BARG_SUB]![SD21]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6019
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =27
                    BackColor =13434828
                    Name ="Text42"
                    ControlSource ="=[KAR_BARG_SUB]![SD22]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5524
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =28
                    BackColor =13434828
                    Name ="Text43"
                    ControlSource ="=[KAR_BARG_SUB]![SD23]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5029
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =29
                    BackColor =13434828
                    Name ="Text44"
                    ControlSource ="=[KAR_BARG_SUB]![SD24]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4534
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =30
                    BackColor =13434828
                    Name ="Text45"
                    ControlSource ="=[KAR_BARG_SUB]![SD25]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4039
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =31
                    BackColor =13434828
                    Name ="Text46"
                    ControlSource ="=[KAR_BARG_SUB]![SD26]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3544
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =32
                    BackColor =13434828
                    Name ="Text47"
                    ControlSource ="=[KAR_BARG_SUB]![SD27]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3049
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =33
                    BackColor =13434828
                    Name ="Text48"
                    ControlSource ="=[KAR_BARG_SUB]![SD28]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2554
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =34
                    BackColor =13434828
                    Name ="Text49"
                    ControlSource ="=[KAR_BARG_SUB]![SD29]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2059
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =35
                    BackColor =13434828
                    Name ="Text50"
                    ControlSource ="=[KAR_BARG_SUB]![SD30]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1564
                    Top =7920
                    Width =449
                    Height =284
                    TabIndex =36
                    BackColor =13434828
                    Name ="Text51"
                    ControlSource ="=[KAR_BARG_SUB]![SD31]"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =7927
                    Width =1484
                    Height =284
                    TabIndex =37
                    BackColor =13434828
                    Name ="Text54"
                    ControlSource ="=[KAR_BARG_SUB]![SZD]"
                End
            End
        End
    End
End
CodeBehindForm
' See "KAR_HEAD.cls"
