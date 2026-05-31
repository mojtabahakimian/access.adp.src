Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14513
    DatasheetFontHeight =10
    ItemSuffix =115
    Left =300
    Top =390
    Right =14820
    Bottom =8265
    HelpContextId =6002
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Filter ="((NUMBER=1696))"
    OrderBy ="UP_TIME DESC"
    RecSrcDt = Begin
        0x8844033552d7e340
    End
    RecordSource ="SELECT TR_HEAD_LST.* FROM TR_HEAD_LST WHERE (TAG = 1) ORDER BY NUMBER"
    Caption ="رسيد انبار"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9127
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12019
                    Top =236
                    Width =1071
                    Height =315
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13234
                            Top =241
                            Width =1110
                            Height =300
                            Name ="Label1"
                            Caption ="شماره رسيد:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10005
                    Top =225
                    Width =996
                    Height =315
                    TabIndex =1
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =11010
                            Top =225
                            Width =900
                            Height =300
                            Name ="Label9"
                            Caption ="تاريخ رسيد:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =210
                    Top =1425
                    Width =14160
                    Height =6855
                    TabIndex =5
                    Name ="INVO_LST_RASID_SUB"
                    SourceObject ="Form.TR_INVO_LST_RASID_SUB"
                    LinkChildFields ="NUMBER;TAG;UP_DATE;UP_TIME"
                    LinkMasterFields ="NUMBER;htag;UP_DATE;UP_TIME"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5494
                    Width =291
                    Height =315
                    TabIndex =6
                    Name ="hTAG"
                    ControlSource ="TAG"
                    DefaultValue ="1"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =737
                    Top =8367
                    Width =966
                    Height =315
                    FontWeight =700
                    TabIndex =7
                    BackColor =13434828
                    Name ="Text59"
                    ControlSource ="=[INVO_LST_RASID_SUB].[Form]![SMABLK]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1802
                            Top =8367
                            Width =1365
                            Height =300
                            FontWeight =700
                            Name ="Label61"
                            Caption ="جمع تحويلي:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4225
                    Top =8372
                    Width =1881
                    Height =315
                    FontWeight =700
                    TabIndex =8
                    BackColor =16777164
                    Name ="MOGU"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6205
                            Top =8372
                            Width =750
                            Height =300
                            FontWeight =700
                            Name ="Label63"
                            Caption ="موجودي:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7479
                    TabIndex =9
                    Name ="N_S"
                    ControlSource ="N_S"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10042
                    Top =615
                    Width =3066
                    Height =315
                    TabIndex =3
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="TAH"
                    ControlSource ="TAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TAH FROM HEAD_LST GROUP BY TAH ORDER BY TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13249
                            Top =615
                            Width =1170
                            Height =300
                            Name ="Label105"
                            Caption ="تحويل گيرنده:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5497
                    Top =634
                    Width =3150
                    Height =315
                    TabIndex =4
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MOLAH FROM HEAD_LST GROUP BY MOLAH ORDER BY MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8797
                            Top =630
                            Width =1215
                            Height =300
                            Name ="Label108"
                            Caption ="تحويل دهنده:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =255
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8073
                    Top =228
                    Width =1311
                    Height =315
                    TabIndex =10
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="uCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =9441
                            Top =225
                            Width =435
                            Height =240
                            ForeColor =8388608
                            Name ="Label111"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5910
                    Top =226
                    Width =1011
                    Height =315
                    TabIndex =2
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6942
                            Top =226
                            Width =1065
                            Height =300
                            Name ="Label103"
                            Caption ="شماره داخلي:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10367
                    Top =8334
                    TabIndex =11
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =10628
                            Top =8277
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10367
                    Top =8588
                    TabIndex =12
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10621
                            Top =8532
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="كنترل كيفيت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10367
                    Top =8847
                    TabIndex =13
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10621
                            Top =8784
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =8272
                    Width =1716
                    Height =315
                    TabIndex =14
                    ForeColor =8388608
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPART.DEPATMAN, DEPART.DEPNAME FROM DEPART ORDER BY DEPART.DEPNAME;"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![DEFAULT]![TFSAZMAN]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =255
                            TextFontCharSet =0
                            Left =10072
                            Width =825
                            Height =300
                            ForeColor =8388608
                            Name ="Label107"
                            Caption ="واحد :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =10144
                    Width =1191
                    Height =315
                    TabIndex =15
                    ForeColor =8388608
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            Left =11374
                            Width =585
                            Height =300
                            ForeColor =8388608
                            Name ="Label109"
                            Caption ="شيفت :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =10771
                    Top =963
                    Width =1566
                    Height =315
                    TabIndex =16
                    ForeColor =8388608
                    Name ="CUST_KIND"
                    ControlSource ="CUST_KIND"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                        "TKNAME;"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    DefaultValue ="[Forms]![BASEKNOW]![DEFTKH]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12376
                            Top =963
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label113"
                            Caption ="نوع مشتري :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3046
                    Width =1191
                    Height =315
                    TabIndex =17
                    ForeColor =8388608
                    Name ="Combo110"
                    ControlSource ="SHIFT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT.SHIFT_ID, SHIFT.SHNAME FROM SHIFT ORDER BY SHIFT.SHNAME;"
                    ColumnWidths ="0"
                    DefaultValue ="=[Forms]![DEFAULT]![SHIFT]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4276
                            Width =585
                            Height =300
                            ForeColor =8388608
                            Name ="Label112"
                            Caption ="شيفت :"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =5328
                    Top =1020
                    Width =1020
                    Height =300
                    ForeColor =8388608
                    Name ="Label114"
                    Caption ="نوع مشتري :"
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =3950
                    Top =677
                    Width =340
                    Height =218
                    TabIndex =18
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4340
                            Top =623
                            Width =795
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =165
                    Top =168
                    Width =1266
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1532
                            Top =165
                            Width =900
                            Height =240
                            Name ="Label148"
                            Caption ="تاريخ اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =182
                    Top =615
                    Width =2676
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2974
                            Top =617
                            Width =1050
                            Height =240
                            Name ="Label149"
                            Caption ="ساعت اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =165
                    Top =1055
                    Width =1536
                    Height =345
                    TabIndex =21
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1752
                            Top =1055
                            Width =1290
                            Height =240
                            Name ="Label150"
                            Caption ="كاربر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2594
                    Top =215
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =22
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4125
                            Top =210
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =3435
                    Top =1005
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =23
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3086
                            Top =1057
                            Width =285
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_HEAD_LST_RASID.cls"
