Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5940
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =2985
    Top =2355
    Right =8925
    Bottom =3300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf6a4399fb34de340
    End
    RecordSource ="SELECT Sum(([MEGHk]-[MEGH_MAR])*GHEYMATKH.MABL) AS KHM, Sum([MABL_K]) AS frm, [f"
        "rm]-[khm] AS sud FROM INVO_LST INNER JOIN GHEYMATKH ON INVO_LST.CODE=GHEYMATKH.C"
        "ODE WHERE (((INVO_LST.NUMBER)=Forms!HEAD_LST_FROOSH2!NUMBER) And ((INVO_LST.TAG)"
        "=2)); "
    Caption ="سود وزيان فاكتور جاري"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =963
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =56
                    Width =4821
                    Height =765
                    FontSize =20
                    Name ="Text0"
                    ControlSource ="=IIf(Forms!HEAD_LST_FROOSH2!VAS=1 Or IsNull(Forms!HEAD_LST_FROOSH2!VAS),[sud]-Fo"
                        "rms!HEAD_LST_FROOSH2!NTKHFIF+Forms!HEAD_LST_FROOSH2!HKH,[sud]-Forms!HEAD_LST_FRO"
                        "OSH2!HKH-Forms!HEAD_LST_FROOSH2!NTKHFIF)"
                    Format ="Standard"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =4920
                            Top =60
                            Width =1020
                            Height =750
                            FontSize =20
                            Name ="Label1"
                            Caption ="سود  :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1984
                    Top =453
                    TabIndex =1
                    Name ="sud"
                    ControlSource ="sud"
                End
            End
        End
    End
End
CodeBehindForm
' See "SOUD2.cls"
