Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14166
    RowHeight =330
    DatasheetFontHeight =10
    ItemSuffix =40
    Left =345
    Top =2760
    Right =8760
    Bottom =8550
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    Tag ="@F_MENU_DATE_DT1=FORMS![F_MENU_DATE]![DT1],@F_MENU_DATE_DT2=FORMS![F_MENU_DATE]!"
        "[DT2]"
    ShortcutMenuBar ="MNEMUDAR"
    RecSrcDt = Begin
        0x69f312a66f64e540
    End
    RecordSource ="SELECT SUM(MABL_K) AS mabkk, col1,col2,col3, COUNT(CODE) AS cont FROM KHAREED_MP"
        "(1, 99999999) KHAREED_MP GROUP BY col1,col2,col3"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:DataAxisEmpty/>\015"
        "\012  <x:PivotField>\015\012   <x:Name>موجودي</x:Name>\015\012   <x:SourceName>M"
        "ANDS</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encod"
        "edType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Fixed</x:NumberFormat>"
        "\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>سايز</x:Name>\015"
        "\012   <x:SourceName>SAIZE</x:SourceName>\015\012   <x:Orientation>Row</x:Orient"
        "ation>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareO"
        "rderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>كد انبار</x:Name"
        ">\015\012   <x:SourceName>ANBAR</x:SourceName>\015\012   <x:DataType>Number</x:D"
        "ataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareO"
        "rderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInF"
        "ieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>كا"
        "رتن</x:Name>\015\012   <x:SourceName>MANDFS</x:SourceName>\015\012   <x:DataType"
        ">Number</x:DataType>\015\012   <x:Orientation>Page</x:Orientation>\015\012   <x:"
        "EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Fixed</x:NumberFo"
        "rmat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy"
        ">\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>انبار</x:Na"
        "me>\015\012   <x:SourceName>ANBARN</x:SourceName>\015\012   <x:EncodedType>adVar"
        "WChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compare"
        "OrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>نام واحد</x:Name>\015\012   <x:SourceName"
        ">NAMES</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </"
        "x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012   "
        "<x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015\012"
        "   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:"
        "Name>Sum of كارتن</x:Name>\015\012   <x:PLName>Total2</x:PLName>\015\012   <x:To"
        "talNumber>1</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012"
        "   <x:Position>1</x:Position>\015\012   <x:ParentField>كارتن</x:ParentField>\015"
        "\012   <x:NumberFormat>Fixed</x:NumberFormat>\015\012  </x:PivotField>\015\012  "
        "<x:PLTotal>\015\012   <x:Name>Total1</x:Name>\015\012   <x:Caption>Sum of موجودي"
        "</x:Caption>\015\012   <x:ParentField>موجودي</x:ParentField>\015\012   <x:TotalN"
        "umber>0</x:TotalNumber>\015\012   <x:NumberFormat>Fixed</x:NumberFormat>\015\012"
        "  </x:PLTotal>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015\012   <x:"
        "TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOffset>0"
        "</x:LeftOffset>\015\012   <x:SeqNum>19</x:SeqNum>\015\012  </x:PivotData>\015\012"
        "  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:C"
        "aption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015"
        "\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:Allo"
        "wAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDe"
        "letions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:DataSource>\015\012   "
        "<x:Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012"
        "  <x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:C"
        "ategory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012  </x:Value>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:Data"
        "SourceIndex>\015\012  </x:BoundCharts>\015\012  <x:FormatValue>\015\012   <x:Dat"
        "aSourceIndex>-3</x:DataSourceIndex>\015\012   <x:Data>2</x:Data>\015\012  </x:Fo"
        "rmatValue>\015\012  <x:PivotAggOrientation>Series</x:PivotAggOrientation>\015\012"
        "  <x:HidePivotFields/>\015\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012"
        "  <x:Palette>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x"
        ":Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry"
        ">\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Entry>\015"
        "\012   <x:Entry>#802060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015\012  "
        " <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012   <x:En"
        "try>#FF8080</x:Entry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:Entry>#C"
        "0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#FF00FF<"
        "/x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF</x:Ent"
        "ry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Entry>\015"
        "\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015\012  "
        "</x:Palette>\015\012  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Border>\015"
        "\012   <x:Color>#2F4F4F</x:Color>\015\012  </x:Border>\015\012  <x:Interior>\015"
        "\012   <x:Color>#FFFFFF</x:Color>\015\012   <x:FillEffect>\015\012    <x:fill x:"
        "type=\"Solid\" x:color=\"#FFFFFF\"/>\015\012   </x:FillEffect>\015\012  </x:Inte"
        "rior>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012    "
        " <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012   "
        "  <x:SubType>3D</x:SubType>\015\012     <x:Series>\015\012      <x:FormatMap>\015"
        "\012      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Ca"
        "ption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:"
        "Data>&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index"
        ">0</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:Dat"
        "aSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x"
        ":DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:F"
        "ormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012     "
        "  <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:DataLabels>\015"
        "\012       <x:Border>\015\012        <x:ColorIndex>None</x:ColorIndex>\015\012  "
        "     </x:Border>\015\012       <x:Font>\015\012        <x:Size>10</x:Size>\015\012"
        "        <x:Color>#FF0000</x:Color>\015\012        <x:B/>\015\012        <x:I>Aut"
        "omatic</x:I>\015\012        <x:U>Automatic</x:U>\015\012       </x:Font>\015\012"
        "       <x:Number>\015\012        <x:FormatString>General</x:FormatString>\015\012"
        "       </x:Number>\015\012       <x:ShowValue/>\015\012       <x:Separator>;</x:"
        "Separator>\015\012       <x:Position>OutsideEnd</x:Position>\015\012      </x:Da"
        "taLabels>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012"
        "      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thicknes"
        "s>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      <x:Identifier>!.Sum of كارتن</x:Identifier>\015\012     </x:Series>\015\012"
        "     <x:Dimension>\015\012      <x:ScaleID>381032116</x:ScaleID>\015\012      <x"
        ":Index>Categories</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>"
        "\015\012      <x:ScaleID>381075476</x:ScaleID>\015\012      <x:Index>Value</x:In"
        "dex>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleI"
        "D>381073788</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015\012     "
        "</x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>381077592</x:Sc"
        "aleID>\015\012      <x:Index>Series</x:Index>\015\012     </x:Dimension>\015\012"
        "     <x:Overlap>0</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012  "
        "   <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:"
        "Axis>\015\012     <x:AxisID>381069684</x:AxisID>\015\012     <x:ScaleID>38103211"
        "6</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012     <x:MajorTick>Out"
        "side</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Pl"
        "acement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015"
        "\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>381045780</x:AxisID>"
        "\015\012     <x:ScaleID>381075476</x:ScaleID>\015\012     <x:Type>Value</x:Type>"
        "\015\012     <x:Number>\015\012      <x:FormatString>Fixed</x:FormatString>\015\012"
        "     </x:Number>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridlines>\015"
        "\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:Min"
        "orTick>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>\015\012"
        "   </x:PlotArea>\015\012   <x:View3D>\015\012    <x:GapDepth>150</x:GapDepth>\015"
        "\012    <x:Perspective>20.0</x:Perspective>\015\012    <x:Rotation>20.0</x:Rotat"
        "ion>\015\012    <x:Inclination>15.0</x:Inclination>\015\012    <x:Light>\015\012"
        "     <x:Rotation>315.0</x:Rotation>\015\012     <x:Inclination>15.0</x:Inclinati"
        "on>\015\012     <x:IntensityDiffuse>0.549019607843137</x:IntensityDiffuse>\015\012"
        "     <x:IntensityAmbient>0.619607843137255</x:IntensityAmbient>\015\012     <x:N"
        "ormal>0.5</x:Normal>\015\012    </x:Light>\015\012   </x:View3D>\015\012   <x:Id"
        "entifier></x:Identifier>\015\012   <x:Walls>\015\012    <x:Index>0</x:Index>\015"
        "\012    <x:Thickness>6</x:Thickness>\015\012   </x:Walls>\015\012   <x:Walls>\015"
        "\012    <x:Index>1</x:Index>\015\012    <x:Thickness>6</x:Thickness>\015\012   <"
        "/x:Walls>\015\012   <x:Walls>\015\012    <x:Index>2</x:Index>\015\012    <x:Thic"
        "kness>6</x:Thickness>\015\012   </x:Walls>\015\012  </x:Chart>\015\012  <x:Title"
        ">\015\012   <x:Font>\015\012    <x:FontName>Sina</x:FontName>\015\012    <x:Colo"
        "r>#228B22</x:Color>\015\012    <x:B>Automatic</x:B>\015\012    <x:I>Automatic</x"
        ":I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012   <x:Border>\015"
        "\012    <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>\015\012   <x:Cap"
        "tion>\015\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012    <x:Data>&q"
        "uot;نمودار موجودي كالا ها&quot;</x:Data>\015\012   </x:Caption>\015\012  </x:Tit"
        "le>\015\012  <x:Legend>\015\012   <x:LegendEntry>\015\012    <x:Index>1</x:Index"
        ">\015\012    <x:Delete/>\015\012   </x:LegendEntry>\015\012   <x:Placement>Right"
        "</x:Placement>\015\012  </x:Legend>\015\012  <x:Scaling>\015\012   <x:ScaleID>38"
        "1032116</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:Scal"
        "eID>381075476</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <"
        "x:ScaleID>381073788</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012"
        "   <x:ScaleID>381077592</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiv"
        "eAlert/>\015\012  <x:HideScreenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =13128
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =9557
                    Top =399
                    ColumnWidth =930
                    ColumnOrder =3
                    Name ="cont"
                    ControlSource ="cont"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12614
                            Top =396
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="تعداد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =9554
                    Top =741
                    ColumnWidth =2790
                    ColumnOrder =4
                    TabIndex =1
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x0000000099ccff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12311
                            Top =741
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="مبلغ خريد"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9579
                    Top =1530
                    ColumnWidth =2940
                    ColumnOrder =2
                    TabIndex =2
                    Name ="col3"
                    ControlSource ="col3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP =3)"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12336
                            Top =1530
                            Width =645
                            Height =240
                            Name ="col3l"
                            Caption ="سايز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9579
                    Top =1890
                    ColumnOrder =5
                    TabIndex =3
                    Name ="col1"
                    ControlSource ="col1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12336
                            Top =1890
                            Width =645
                            Height =240
                            Name ="col1l"
                            Caption ="سايز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9576
                    Top =2325
                    ColumnWidth =735
                    ColumnOrder =0
                    TabIndex =4
                    Name ="col2"
                    ControlSource ="col2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12333
                            Top =2325
                            Width =645
                            Height =240
                            Name ="col2l2"
                            Caption ="سايز"
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
                    Left =9576
                    Top =2685
                    ColumnWidth =810
                    ColumnOrder =1
                    TabIndex =5
                    Name ="col31"
                    ControlSource ="col3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12333
                            Top =2685
                            Width =645
                            Height =240
                            Name ="col3l2"
                            Caption ="سايز"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Top =4308
                    Width =8445
                    Height =8820
                    TabIndex =6
                    Name ="khareed_mp4"
                    SourceObject ="Form.khareed_mp4"
                    LinkChildFields ="col1;col2;col3"
                    LinkMasterFields ="col1;col2;col3"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7515
                            Top =4068
                            Width =930
                            Height =240
                            Name ="froosh_mp4 Label"
                            Caption ="froosh_mp4"
                            EventProcPrefix ="froosh_mp4_Label"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "khareed_mp3.cls"
