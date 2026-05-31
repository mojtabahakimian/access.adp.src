Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =1
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14796
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =6825
    Top =3675
    Right =21615
    Bottom =16650
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xb2ce950fb371e340
    End
    Caption ="عيدي پاداش و حق سنوات"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin FormHeader
            SpecialEffect =2
            Height =1332
            BackColor =16777164
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =680
                    Top =56
                    Width =1986
                    Height =576
                    FontWeight =700
                    ForeColor =16711680
                    Name ="Command1"
                    Caption ="پردازش و محاسبه عيدي"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    Tag ="نحوه محاسبه عيدي:\015\01260 روز دستمزد روزانه به شرط آنكه از 90 روز حداقل دستمزد"
                        " بيشتر نشود براي 12 ماه كاركرد\015\012در صورتيكه كاركرد كمتر از 12 ماه باشد به ن"
                        "سبت روزهاي كاركرد محاسبه ميشود\015\012محاسبه ماليات عيدي:\015\012درامد مشمول مال"
                        "يات كه شامل درآمد كل سال طبق ليست ارسالي به اضافه عيدي محاسبه شده  ميباشد \015\012"
                        "منهاي معافيت سالانه (13* حداقل حقوق به نسبت كاركرد) *10% منهاي كل ماليات پرداختي"
                        " سال\015\012 محاسبه روزهاي كاركرد:\015\012روزهاي كاركرد رسمي ثبت شده در قسمت حقو"
                        "ق و دستمزد   به اضافه روزهاي مرخصي پزشكي\015\012(مرخصي زايمان و ...)\015\012نحوه"
                        " محاسبه حق سنوات :\015\012معادل 30 روز حقوق  به ازاي 12 ماه كاركرد\015\012در صور"
                        "تيكه كاركرد كمتر از 12 ماه باشد به نسبت روزهاي كاركرد محاسبه ميشود"
                    ControlTipText ="نحوه محاسبه عيدي:\015\01260 روز دستمزد روزانه به شرط آنكه از 90 روز حداقل دستمزد"
                        " بيشتر نشود براي 12 ماه كاركرد\015\012در صورتيكه كاركرد كمتر از 12 ماه باشد به ن"
                        "سبت روزهاي كاركرد محاسبه ميشود"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10200
                    Top =60
                    Width =1418
                    Height =280
                    TabIndex =1
                    Name ="MINSAL"
                    Format ="#,###"
                    DefaultValue ="73200"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =11681
                            Top =64
                            Width =2925
                            Height =240
                            Name ="Label4"
                            Caption ="حداقل دستمزد طبق قانون كار را وارد كنيد:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10196
                    Top =456
                    Width =1418
                    Height =280
                    TabIndex =2
                    Name ="MAXEYDY"
                    ControlSource ="=[MINSAL]*90"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =11673
                            Top =450
                            Width =1530
                            Height =300
                            Name ="Label6"
                            Caption ="حد اكثر عيدي پاداش:"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =680
                    Top =630
                    Width =1986
                    Height =576
                    FontWeight =700
                    TabIndex =3
                    Name ="Command9"
                    Caption ="چاپ ليست عيدي پاداش"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="چاپ ليست عيدي پاداش"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2685
                    Top =60
                    Width =2196
                    Height =576
                    FontWeight =700
                    TabIndex =4
                    ForeColor =0
                    Name ="Command10"
                    Caption ="مشاهده ليست عيدي پاداش محاسبه شده قبلي"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="مشاهده ليست عيدي پاداش محاسبه شده قبلي"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2686
                    Top =630
                    Width =2196
                    Height =576
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="Command11"
                    Caption ="چاپ ليست حق سنوات عيدي و پاداش"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="چاپ ليست حق سنوات عيدي و پاداش"
                End
            End
        End
        Begin Section
            Visible = NotDefault
            CanGrow = NotDefault
            Height =7766
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =56
                    Width =14730
                    Height =7590
                    Name ="SALARY_EYDY_SUB"
                    SourceObject ="Form.SALARY_EYDY_SUB"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "EYDY_PADASH.cls"
