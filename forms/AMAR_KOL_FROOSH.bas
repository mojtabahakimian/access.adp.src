Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5726
    RowHeight =375
    ItemSuffix =8
    Left =600
    Top =1320
    Right =8295
    Bottom =5430
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x4666cc308cf0e240
    End
    ServerFilter ="GRP = 2 AND VH = 1"
    RecordSource ="AMAR_KOL_FROOSH"
    Caption ="آمار فروش به تفكيك هر ماه و گروه كالا"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>مقدار</x:Name>\015\012   <x:SourceName>MEGHT</x:SourceName>\015\012  "
        " <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:E"
        "ncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueN"
        "ame</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>"
        "\015\012   <x:Name>گروه كالا</x:Name>\015\012   <x:SourceName>NAMES</x:SourceNam"
        "e>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrder"
        "edMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>واحد</x:Name>\015\012   <x:SourceName>vahed</"
        "x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:C"
        "ompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Pivot"
        "Field>\015\012  <x:PivotField>\015\012   <x:Name>ماه</x:Name>\015\012   <x:Sourc"
        "eName>MON</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x"
        ":Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adInteger</x:EncodedTy"
        "pe>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012   <x:AutoSortOrder>Ascending</x:AutoSortOrder>\015\012   <x:Expanded/>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ماه1</x:Name>\015\012"
        "   <x:SourceName>MN</x:SourceName>\015\012   <x:PLCaption>ماه</x:PLCaption>\015\012"
        "   <x:FilterCaption>ماه</x:FilterCaption>\015\012   <x:Orientation>Row</x:Orient"
        "ation>\015\012   <x:Position>2</x:Position>\015\012   <x:EncodedType>adVarWChar<"
        "/x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdere"
        "dMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotFi"
        "eld>\015\012   <x:Name>GRP</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMember"
        "sBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>VH</x:Name>\015\012"
        "   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedType>adInteger</x:Encode"
        "dType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersB"
        "y>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientation>Column</x:Orien"
        "tation>\015\012   <x:Position>-1</x:Position>\015\012   <x:DataField/>\015\012  "
        "</x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of مقدار</x:Name>\015"
        "\012   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015"
        "\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015"
        "\012   <x:ParentField>مقدار</x:ParentField>\015\012  </x:PivotField>\015\012  <x"
        ":PivotData>\015\012   <x:Top>0.0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset"
        ">\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012 "
        "  <x:SeqNum>10</x:SeqNum>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012"
        "   <x:IsNotFiltered/>\015\012   <x:DetailSortOrder>MON</x:DetailSortOrder>\015\012"
        "   <x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015"
        "\012    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:Al"
        "lowEdits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:Allo"
        "wDeletions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTab"
        "le>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012   <x:Data>Total1</x:Data>\015\012  </"
        "x:Value>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:DataSourceIn"
        "dex>\015\012  </x:BoundCharts>\015\012  <x:FormatValue>\015\012   <x:DataSourceI"
        "ndex>-3</x:DataSourceIndex>\015\012   <x:Data>2</x:Data>\015\012  </x:FormatValu"
        "e>\015\012  <x:PivotAggOrientation>None</x:PivotAggOrientation>\015\012  <x:Hide"
        "PivotFields/>\015\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Pal"
        "ette>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Entry>\015\012  "
        " <x:Entry>#802060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015\012   <x:En"
        "try>#A0E0E0</x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012   <x:Entry>#F"
        "F8080</x:Entry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:Entry>#C0C0FF<"
        "/x:Entry>\015\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#FF00FF</x:Ent"
        "ry>\015\012   <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF</x:Entry>\015"
        "\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Entry>\015\012  "
        " <x:Entry>#FFC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015\012  </x:Pa"
        "lette>\015\012  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Border>\015\012"
        "   <x:Color>#2F4F4F</x:Color>\015\012  </x:Border>\015\012  <x:Interior>\015\012"
        "   <x:Color>#EEE8AA</x:Color>\015\012   <x:FillEffect>\015\012    <x:fill x:type"
        "=\"Solid\" x:color=\"#EEE8AA\"/>\015\012   </x:FillEffect>\015\012  </x:Interior"
        ">\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012     <x:"
        "SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012     <x"
        ":Series>\015\012      <x:FormatMap>\015\012       <x:Segment>\015\012        <x:"
        "SegmentBoundary>\015\012         <x:Index>0</x:Index>\015\012         <x:Interio"
        "r>\015\012          <x:Color>#FF0000</x:Color>\015\012         </x:Interior>\015"
        "\012         <x:Percentage/>\015\012         <x:Value>0.0</x:Value>\015\012     "
        "   </x:SegmentBoundary>\015\012        <x:SegmentBoundary>\015\012         <x:In"
        "dex>1</x:Index>\015\012         <x:Interior>\015\012          <x:Color>#0000FF</"
        "x:Color>\015\012         </x:Interior>\015\012         <x:Percentage/>\015\012  "
        "       <x:Value>1.0</x:Value>\015\012        </x:SegmentBoundary>\015\012       "
        " <x:HasAbsoluteLabels/>\015\012       </x:Segment>\015\012      </x:FormatMap>\015"
        "\012      <x:Name>Series</x:Name>\015\012      <x:Caption>\015\012       <x:Data"
        "SourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Series&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012       <x:Data>Total1</x:Data>\015\012      </x:Value>\015\012"
        "      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015"
        "\012       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Data"
        "Labels>\015\012       <x:Border>\015\012        <x:ColorIndex>None</x:ColorIndex"
        ">\015\012       </x:Border>\015\012       <x:Number>\015\012        <x:FormatStr"
        "ing>General</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowValu"
        "e/>\015\012       <x:Separator>;</x:Separator>\015\012       <x:Position>Outside"
        "End</x:Position>\015\012      </x:DataLabels>\015\012      <x:Marker>\015\012   "
        "    <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0"
        "</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSour"
        "ceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!</x:Identifier>\015\012"
        "     </x:Series>\015\012     <x:Dimension>\015\012      <x:ScaleID>492333688</x:"
        "ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     </x:Dimension>\015"
        "\012     <x:Dimension>\015\012      <x:ScaleID>492359384</x:ScaleID>\015\012    "
        "  <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015"
        "\012      <x:ScaleID>492359588</x:ScaleID>\015\012      <x:Index>FormatValue</x:"
        "Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012  "
        "   <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAn"
        "gle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>492310388</"
        "x:AxisID>\015\012     <x:ScaleID>492333688</x:ScaleID>\015\012     <x:Type>Categ"
        "ory</x:Type>\015\012     <x:Font>\015\012      <x:FontName>Tahoma</x:FontName>\015"
        "\012      <x:Size>8</x:Size>\015\012      <x:B>Automatic</x:B>\015\012      <x:I"
        ">Automatic</x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x:Font>\015\012"
        "     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTi"
        "ck>\015\012     <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Au"
        "to</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:Axis"
        "ID>492311096</x:AxisID>\015\012     <x:ScaleID>492359384</x:ScaleID>\015\012    "
        " <x:Type>Value</x:Type>\015\012     <x:Line>\015\012      <x:Weight>0</x:Weight>"
        "\015\012     </x:Line>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridl"
        "ines>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>No"
        "ne</x:MinorTick>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:Fo"
        "ntName>Tahoma</x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Fo"
        "nt>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012"
        "      </x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</"
        "x:DataSourceIndex>\015\012       <x:Data>&quot;ميزان فروش&quot;</x:Data>\015\012"
        "      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012     </x:Ti"
        "tle>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>\015\012   "
        "</x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Chart>\015\012"
        "  <x:Title>\015\012   <x:Font>\015\012    <x:FontName>Sina</x:FontName>\015\012 "
        "   <x:Size>11</x:Size>\015\012    <x:Color>#A52A2A</x:Color>\015\012    <x:B>Aut"
        "omatic</x:B>\015\012    <x:I>Automatic</x:I>\015\012    <x:U>Automatic</x:U>\015"
        "\012   </x:Font>\015\012   <x:Border>\015\012    <x:ColorIndex>None</x:ColorInde"
        "x>\015\012   </x:Border>\015\012   <x:Caption>\015\012    <x:DataSourceIndex>-1<"
        "/x:DataSourceIndex>\015\012    <x:Data>&quot;نمودار ميزان فروش تعداي محصولات در "
        "هر ماه&quot;</x:Data>\015\012   </x:Caption>\015\012  </x:Title>\015\012  <x:Sca"
        "ling>\015\012   <x:ScaleID>492333688</x:ScaleID>\015\012  </x:Scaling>\015\012  "
        "<x:Scaling>\015\012   <x:ScaleID>492359384</x:ScaleID>\015\012  </x:Scaling>\015"
        "\012  <x:Scaling>\015\012   <x:ScaleID>492359588</x:ScaleID>\015\012  </x:Scalin"
        "g>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideScreenTips/>\015\012 </x:Chart"
        "Space>\015\012</xml>"
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
            Height =2734
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
                    Left =2551
                    Top =453
                    ColumnOrder =2
                    Name ="MEGHT"
                    ControlSource ="MEGHT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =453
                            Width =645
                            Height =240
                            Name ="Label0"
                            Caption ="مقدار"
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
                    Left =2551
                    Top =793
                    ColumnWidth =1815
                    ColumnOrder =3
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =793
                            Width =645
                            Height =240
                            Name ="Label1"
                            Caption ="گروه كالا"
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
                    Left =2551
                    Top =1133
                    ColumnOrder =4
                    TabIndex =2
                    Name ="vahed"
                    ControlSource ="vahed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1133
                            Width =585
                            Height =240
                            Name ="Label2"
                            Caption ="واحد"
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
                    Left =2551
                    Top =1474
                    ColumnWidth =390
                    ColumnOrder =0
                    TabIndex =3
                    Name ="MON"
                    ControlSource ="MON"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1474
                            Width =480
                            Height =240
                            Name ="Label3"
                            Caption ="ماه"
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
                    Left =2551
                    Top =1834
                    ColumnOrder =1
                    TabIndex =4
                    Name ="MN"
                    ControlSource ="MN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1834
                            Width =480
                            Height =240
                            Name ="Label5"
                            Caption ="ماه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AMAR_KOL_FROOSH.cls"
