Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11338
    RowHeight =345
    ItemSuffix =9
    Left =495
    Top =990
    Right =11880
    Bottom =5070
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Tag ="@FROMDATE=FORMS![F_MENU_DATE]![DT1],@TODATE=FORMS![F_MENU_DATE]![DT2]"
    RecSrcDt = Begin
        0x5897f1921960e640
    End
    RecordSource ="SELECT INVO_LST.CODE, SUM(INVO_LST.MEGHk) AS MEGHT, STUF_DEF.NAME, TCOD_STUFGROU"
        "P.NAMES, TCOD_VAHEDS.NAMES AS vahed, STUF_DEF.N_FANI FROM HEAD_LST INNER JOIN IN"
        "VO_LST ON HEAD_LST.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG = INVO_LST.TAG INNE"
        "R JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE INNER JOIN TCOD_STUFGROUP ON ST"
        "UF_DEF.RADAH = TCOD_STUFGROUP.CODE INNER JOIN TCOD_VAHEDS ON STUF_DEF.VAHED = TC"
        "OD_VAHEDS.CODE WHERE (HEAD_LST.TAG = 9) AND (HEAD_LST.DEPATMAN = 200) GROUP BY I"
        "NVO_LST.CODE, STUF_DEF.NAME, TCOD_STUFGROUP.NAMES, TCOD_VAHEDS.NAMES, STUF_DEF.N"
        "_FANI"
    Caption ="آمار توليد"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:Cac"
        "heDetails/>\015\012  <x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  "
        "<x:PivotField>\015\012   <x:Name>كد كالا</x:Name>\015\012   <x:SourceName>CODE</"
        "x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:C"
        "ompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Pivot"
        "Field>\015\012  <x:PivotField>\015\012   <x:Name>مقدار توليد</x:Name>\015\012   "
        "<x:SourceName>MEGHT</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام كالال</x:Name>\015"
        "\012   <x:SourceName>NAME</x:SourceName>\015\012   <x:Orientation>Row</x:Orienta"
        "tion>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOr"
        "deredMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:AutoSortOrder"
        ">Ascending</x:AutoSortOrder>\015\012   <x:AutoSortField>Sum of مقدار توليد</x:Au"
        "toSortField>\015\012   <x:FilterFunctionValue>1</x:FilterFunctionValue>\015\012 "
        "  <x:FilterOn>Total1</x:FilterOn>\015\012   <x:Expanded/>\015\012  </x:PivotFiel"
        "d>\015\012  <x:PivotField>\015\012   <x:Name>گروه كالا</x:Name>\015\012   <x:Sou"
        "rceName>NAMES</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>"
        "\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام واحد</x:Name"
        ">\015\012   <x:SourceName>vahed</x:SourceName>\015\012   <x:Orientation>Page</x:"
        "Orientation>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Co"
        "mpareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotF"
        "ield>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012   <x:Orien"
        "tation>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015\012   <x"
        ":DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>"
        "Sum of مقدار توليد</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012   <x:T"
        "otalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012"
        "   <x:Position>1</x:Position>\015\012   <x:ParentField>مقدار توليد</x:ParentFiel"
        "d>\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015"
        "\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:Le"
        "ftOffset>0</x:LeftOffset>\015\012   <x:InvertedRowMember>!.30*60اولاس درجه سه رد"
        "ه E كد 13  B24</x:InvertedRowMember>\015\012   <x:SeqNum>114</x:SeqNum>\015\012 "
        " </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x"
        ":DetailSortOrder>NAME</x:DetailSortOrder>\015\012   <x:Label>\015\012    <x:Capt"
        "ion>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015\012"
        "   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:AllowAdd"
        "itions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDeleti"
        "ons>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>14552</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Close>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012   <x:Data>Total1</x:Data>\015\012  </"
        "x:Close>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:DataSourceIn"
        "dex>\015\012  </x:BoundCharts>\015\012  <x:PivotAggOrientation>None</x:PivotAggO"
        "rientation>\015\012  <x:HidePivotFields/>\015\012  <x:UnifiedScales/>\015\012  <"
        "x:Palette>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:En"
        "try>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Entry>\015"
        "\012   <x:Entry>#802060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015\012  "
        " <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012   <x:En"
        "try>#FF8080</x:Entry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:Entry>#C"
        "0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#FF00FF<"
        "/x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF</x:Ent"
        "ry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Entry>\015"
        "\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015\012  "
        "</x:Palette>\015\012  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Border>\015"
        "\012   <x:Color>#FF1493</x:Color>\015\012  </x:Border>\015\012  <x:Interior>\015"
        "\012   <x:Color>#8FBC8F</x:Color>\015\012   <x:FillEffect>\015\012    <x:fill x:"
        "type=\"Solid\" x:color=\"#8FBC8F\"/>\015\012   </x:FillEffect>\015\012  </x:Inte"
        "rior>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012    "
        " <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012   "
        "  <x:SubType>3D</x:SubType>\015\012     <x:Series>\015\012      <x:FormatMap>\015"
        "\012      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Ca"
        "ption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:"
        "Data>&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index"
        ">0</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:Dat"
        "aSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x"
        ":DataSourceIndex>0</x:DataSourceIndex>\015\012       <x:Data>Total1</x:Data>\015"
        "\012      </x:Value>\015\012      <x:DataLabels>\015\012       <x:Border>\015\012"
        "        <x:ColorIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012     "
        "  <x:Number>\015\012        <x:FormatString>General</x:FormatString>\015\012    "
        "   </x:Number>\015\012       <x:ShowValue/>\015\012       <x:Separator>;</x:Sepa"
        "rator>\015\012       <x:Position>OutsideEnd</x:Position>\015\012      </x:DataLa"
        "bels>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012   "
        "   </x:Marker>\015\012      <x:Explode>10</x:Explode>\015\012      <x:Thickness>"
        "10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  "
        "    <x:Identifier>!</x:Identifier>\015\012     </x:Series>\015\012     <x:Dimens"
        "ion>\015\012      <x:ScaleID>372717728</x:ScaleID>\015\012      <x:Index>Categor"
        "ies</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      "
        "<x:ScaleID>372717932</x:ScaleID>\015\012      <x:Index>Value</x:Index>\015\012  "
        "   </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>372718136</x"
        ":ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015\012     </x:Dimension>"
        "\015\012     <x:Dimension>\015\012      <x:ScaleID>372718340</x:ScaleID>\015\012"
        "      <x:Index>Series</x:Index>\015\012     </x:Dimension>\015\012     <x:Overla"
        "p>0</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSli"
        "ceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012 "
        "    <x:AxisID>372721492</x:AxisID>\015\012     <x:ScaleID>372717932</x:ScaleID>\015"
        "\012     <x:Type>Value</x:Type>\015\012     <x:Number>\015\012      <x:FormatStr"
        "ing>General</x:FormatString>\015\012     </x:Number>\015\012     <x:MajorGridlin"
        "es>\015\012     </x:MajorGridlines>\015\012     <x:MajorTick>Outside</x:MajorTic"
        "k>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Placement>Left</x:"
        "Placement>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>372722"
        "332</x:AxisID>\015\012     <x:ScaleID>372717728</x:ScaleID>\015\012     <x:Type>"
        "Category</x:Type>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:"
        "MinorTick>None</x:MinorTick>\015\012     <x:Placement>Bottom</x:Placement>\015\012"
        "     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axis>\015\012   </x:Pl"
        "otArea>\015\012   <x:View3D>\015\012    <x:GapDepth>150</x:GapDepth>\015\012    "
        "<x:RightAngleAxes/>\015\012    <x:ExtrudeAngle>45.0</x:ExtrudeAngle>\015\012    "
        "<x:Light>\015\012     <x:Rotation>315.0</x:Rotation>\015\012     <x:Inclination>"
        "15.0</x:Inclination>\015\012     <x:IntensityDiffuse>0.549019607843137</x:Intens"
        "ityDiffuse>\015\012     <x:IntensityAmbient>0.619607843137255</x:IntensityAmbien"
        "t>\015\012     <x:Normal>0.5</x:Normal>\015\012    </x:Light>\015\012   </x:View"
        "3D>\015\012   <x:Identifier></x:Identifier>\015\012   <x:Walls>\015\012    <x:In"
        "dex>0</x:Index>\015\012    <x:Thickness>6</x:Thickness>\015\012   </x:Walls>\015"
        "\012   <x:Walls>\015\012    <x:Index>1</x:Index>\015\012    <x:Thickness>6</x:Th"
        "ickness>\015\012   </x:Walls>\015\012   <x:Walls>\015\012    <x:Index>2</x:Index"
        ">\015\012    <x:Thickness>6</x:Thickness>\015\012   </x:Walls>\015\012  </x:Char"
        "t>\015\012  <x:Title>\015\012   <x:Interior>\015\012    <x:ColorIndex>None</x:Co"
        "lorIndex>\015\012    <x:FillEffect>\015\012     <x:fill x:type=\"Solid\" x:color"
        "=\"#000000\"/>\015\012    </x:FillEffect>\015\012   </x:Interior>\015\012   <x:F"
        "ont>\015\012    <x:FontName>Titr</x:FontName>\015\012    <x:B>Automatic</x:B>\015"
        "\012    <x:I>Automatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>"
        "\015\012   <x:Border>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012    <x"
        ":Weight>3</x:Weight>\015\012   </x:Border>\015\012   <x:Caption>\015\012    <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012    <x:Data>&quot;نمودار توليد&quot"
        ";</x:Data>\015\012   </x:Caption>\015\012  </x:Title>\015\012  <x:Scaling>\015\012"
        "   <x:ScaleID>372717728</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015"
        "\012   <x:ScaleID>372717932</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scalin"
        "g>\015\012   <x:ScaleID>372718136</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:"
        "Scaling>\015\012   <x:ScaleID>372718340</x:ScaleID>\015\012  </x:Scaling>\015\012"
        "  <x:HasSelectionFeedback>True</x:HasSelectionFeedback>\015\012  <x:HidePassiveA"
        "lert/>\015\012 </x:ChartSpace>\015\012</xml>"
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
            Height =1340
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
                    Left =8562
                    Top =56
                    Width =2490
                    ColumnWidth =960
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6748
                            Top =56
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا*"
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
                    Left =8562
                    Top =396
                    ColumnWidth =1260
                    TabIndex =1
                    Name ="MEGHT"
                    ControlSource ="MEGHT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6755
                            Top =393
                            Width =810
                            Height =240
                            Name ="Label3"
                            Caption ="مقدار توليد"
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
                    Left =8562
                    Top =736
                    Width =2490
                    ColumnWidth =4350
                    TabIndex =2
                    Name ="NAMEe"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6755
                            Top =738
                            Width =645
                            Height =240
                            Name ="Label5"
                            Caption ="نام كالال"
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
                    Left =8562
                    Top =1098
                    Width =2490
                    ColumnWidth =1575
                    TabIndex =3
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6755
                            Top =1100
                            Width =645
                            Height =240
                            Name ="Label7"
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
                    Left =1473
                    Top =396
                    TabIndex =4
                    Name ="vahed"
                    ControlSource ="vahed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4185
                            Top =396
                            Width =690
                            Height =240
                            Name ="Label8"
                            Caption ="نام واحد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AMAR_TOLID.cls"
