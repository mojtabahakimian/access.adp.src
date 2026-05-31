Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7597
    RowHeight =330
    ItemSuffix =21
    Left =645
    Top =375
    Right =14700
    Bottom =7875
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x63db54986fe7e340
    End
    RecordSource ="SELECT SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF) AS MEGHF, SUM(FROOSH_COUNTALL_CIT"
        "Y_N_FANI.MABK) AS MABKK, SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHkK) AS MEGHk, FROOS"
        "H_COUNTALL_CITY_N_FANI.OSCO, FROOSH_COUNTALL_CITY_N_FANI.CUST_NO, FROOSH_COUNTAL"
        "L_CITY_N_FANI.CUST_NO AS CUSTNO, TCOD_OSTAN.OSNAME, CUST_HESAB.NAME FROM FROOSH_"
        "COUNTALL_CITY_N_FANI(1, 999999) FROOSH_COUNTALL_CITY_N_FANI INNER JOIN TCOD_OSTA"
        "N ON FROOSH_COUNTALL_CITY_N_FANI.OSCO = TCOD_OSTAN.OSCODE INNER JOIN CUST_HESAB "
        "ON FROOSH_COUNTALL_CITY_N_FANI.CUST_NO = CUST_HESAB.hes GROUP BY FROOSH_COUNTALL"
        "_CITY_N_FANI.OSCO, FROOSH_COUNTALL_CITY_N_FANI.CUST_NO, TCOD_OSTAN.OSNAME, CUST_"
        "HESAB.NAME ORDER BY SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF)"
    Caption ="ليست فروش روزانه به تفكيك نوع كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>كارتن</x:Name>\015\012   <x:SourceName>MEGHF</x:SourceName>\015\012  "
        " <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:E"
        "ncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueN"
        "ame</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>"
        "\015\012   <x:Name>مبلغ</x:Name>\015\012   <x:SourceName>MABKK</x:SourceName>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>متراژ</x"
        ":Name>\015\012   <x:SourceName>MEGHk</x:SourceName>\015\012   <x:DataType>Number"
        "</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Com"
        "pareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFi"
        "eld>\015\012  <x:PivotField>\015\012   <x:Name>استان</x:Name>\015\012   <x:Sourc"
        "eName>OSCO</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <"
        "x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Uni"
        "queName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012 "
        " </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>كد مشتري</x:Name>\015"
        "\012   <x:SourceName>CUST_NO</x:SourceName>\015\012   <x:Orientation>Row</x:Orie"
        "ntation>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Compar"
        "eOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>CUSTNO</x:Na"
        "me>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrde"
        "redMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFiel"
        "dList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام م"
        "شتري</x:Name>\015\012   <x:SourceName>OSNAME</x:SourceName>\015\012   <x:Encoded"
        "Type>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName<"
        "/x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>نام مشتري1</x:Name>\015\012   <x:SourceName>NAME</x:SourceName>\015"
        "\012   <x:PLCaption>نام مشتري</x:PLCaption>\015\012   <x:FilterCaption>نام مشتري"
        "</x:FilterCaption>\015\012   <x:Orientation>Row</x:Orientation>\015\012   <x:Pos"
        "ition>2</x:Position>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <"
        "x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>"
        "Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Posit"
        "ion>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>Sum of كارتن</x:Name>\015\012   <x:PLName>Total1"
        "</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>"
        "Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x:ParentFie"
        "ld>كارتن</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012"
        "   <x:Top>0.0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left"
        ">0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>14</x:S"
        "eqNum>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered"
        "/>\015\012   <x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x"
        ":Caption>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdit"
        "s>false</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012"
        "   <x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </"
        "x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:DataSource>\015\012   "
        "<x:Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012"
        "  <x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:C"
        "ategory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012   <x:Data>Total1</x:Data>\015\012  </x:Value>\015\012  <x:BoundCharts>\015\012"
        "   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012  <"
        "x:PivotAggOrientation>None</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>"
        "\015\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#8"
        "02060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0<"
        "/x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012   <x:Entry>#FF8080</x:Ent"
        "ry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015"
        "\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012  "
        " <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:En"
        "try>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#F"
        "FC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012"
        "  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>"
        "\015\012    <x:Graph>\015\012     <x:SubType>Clustered</x:SubType>\015\012     <"
        "x:Type>Column</x:Type>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012"
        "      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Captio"
        "n>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data"
        ">&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</"
        "x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012       <x:Data>Total1</x:Data>\015\012"
        "      </x:Value>\015\012      <x:DataLabels>\015\012       <x:Border>\015\012   "
        "     <x:ColorIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012       <"
        "x:Number>\015\012        <x:FormatString>General</x:FormatString>\015\012       "
        "</x:Number>\015\012       <x:ShowValue/>\015\012       <x:Separator>;</x:Separat"
        "or>\015\012       <x:Position>OutsideEnd</x:Position>\015\012      </x:DataLabel"
        "s>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      "
        "</x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</"
        "x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "<x:Identifier>!</x:Identifier>\015\012     </x:Series>\015\012     <x:Dimension>"
        "\015\012      <x:ScaleID>410598580</x:ScaleID>\015\012      <x:Index>Categories<"
        "/x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:S"
        "caleID>410598800</x:ScaleID>\015\012      <x:Index>Value</x:Index>\015\012     <"
        "/x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>410599020</x:Sca"
        "leID>\015\012      <x:Index>FormatValue</x:Index>\015\012     </x:Dimension>\015"
        "\012     <x:Overlap>0</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012"
        "     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <"
        "x:Axis>\015\012     <x:AxisID>410603664</x:AxisID>\015\012     <x:ScaleID>410598"
        "580</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012     <x:MajorTick>O"
        "utside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:"
        "Title>\015\012      <x:Font>\015\012       <x:FontName>Titr</x:FontName>\015\012"
        "       <x:Size>10</x:Size>\015\012       <x:B/>\015\012      </x:Font>\015\012  "
        "    <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </x:"
        "Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSour"
        "ceIndex>\015\012       <x:Data>&quot;مشتري&quot;</x:Data>\015\012      </x:Capti"
        "on>\015\012      <x:Position>Bottom</x:Position>\015\012     </x:Title>\015\012 "
        "    <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:Groupi"
        "ngEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>410604184"
        "</x:AxisID>\015\012     <x:ScaleID>410598800</x:ScaleID>\015\012     <x:Type>Val"
        "ue</x:Type>\015\012     <x:Number>\015\012      <x:FormatString>General</x:Forma"
        "tString>\015\012     </x:Number>\015\012     <x:MajorGridlines>\015\012     </x:"
        "MajorGridlines>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:Mi"
        "norTick>None</x:MinorTick>\015\012     <x:Title>\015\012      <x:Font>\015\012  "
        "     <x:FontName>Titr</x:FontName>\015\012       <x:Size>10</x:Size>\015\012    "
        "   <x:B/>\015\012      </x:Font>\015\012      <x:Border>\015\012       <x:ColorI"
        "ndex>None</x:ColorIndex>\015\012      </x:Border>\015\012      <x:Caption>\015\012"
        "       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;كار"
        "تن&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Position>Left</x:Pos"
        "ition>\015\012     </x:Title>\015\012     <x:Placement>Left</x:Placement>\015\012"
        "    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015"
        "\012  </x:Chart>\015\012  <x:Scaling>\015\012   <x:ScaleID>410598580</x:ScaleID>"
        "\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>410598800</x:Sc"
        "aleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>410599020"
        "</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  <x:Hi"
        "deScreenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
    AllowFormView =0
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
        Begin Section
            CanGrow = NotDefault
            Height =5215
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4591
                    Top =113
                    ColumnOrder =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="متراژ"
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
                    Left =4591
                    Top =1133
                    ColumnWidth =1065
                    ColumnOrder =4
                    TabIndex =2
                    BackColor =13434828
                    Name ="MEGHF"
                    ControlSource ="MEGHF"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =1133
                            Width =960
                            Height =240
                            Name ="Label7"
                            Caption ="كارتن"
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
                    Left =4591
                    Top =1493
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =1493
                            Width =960
                            Height =240
                            Name ="Label15"
                            Caption ="مبلغ"
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
                    Left =3451
                    Top =3393
                    Width =2490
                    ColumnOrder =0
                    TabIndex =5
                    Name ="Combo18"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1648
                            Top =3400
                            Width =870
                            Height =240
                            Name ="Label19"
                            Caption ="كد مشتري"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3860
                    Top =3798
                    TabIndex =6
                    Name ="OSCO"
                    ControlSource ="OSCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =3798
                            Width =705
                            Height =240
                            Name ="Label20"
                            Caption ="كد استان"
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
                    Left =3451
                    Top =3026
                    Width =2490
                    ColumnWidth =3495
                    ColumnOrder =1
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1648
                            Top =3033
                            Width =870
                            Height =240
                            Name ="Label17"
                            Caption ="نام مشتري"
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
                    Left =4591
                    Top =793
                    Width =2490
                    ColumnWidth =1620
                    ColumnOrder =2
                    TabIndex =1
                    Name ="OSNAME"
                    ControlSource ="OSNAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2784
                            Top =795
                            Width =540
                            Height =240
                            Name ="Label5"
                            Caption ="استان"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNTALL_CITY1_OSTAN_CUSTO.cls"
