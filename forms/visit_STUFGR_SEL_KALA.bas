Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6349
    RowHeight =345
    ItemSuffix =6
    Left =4335
    Top =1050
    Right =11160
    Bottom =8310
    HelpContextId =45
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x955fb7161292e440
    End
    RecordSource ="SELECT CODE, NAMES, TIC FROM TCODE_MENUITEM"
    Caption ="تعريف  گروه كالاها"
    OnCurrent ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin Section
            Height =1303
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =4425
                    Top =993
                    TabIndex =1
                    Name ="Check4"
                    ControlSource ="TIC"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3855
                            Top =963
                            Width =600
                            Height =240
                            Name ="Label5"
                            Caption ="انتخاب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1907
                    Top =453
                    Width =3510
                    Height =300
                    ColumnWidth =4665
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5498
                            Top =450
                            Width =600
                            Height =285
                            Name ="Label3"
                            Caption ="نام گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "visit_STUFGR_SEL_KALA.cls"
