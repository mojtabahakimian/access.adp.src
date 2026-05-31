Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =4
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5726
    RowHeight =375
    ItemSuffix =9
    Left =9930
    Top =3915
    Right =15045
    Bottom =8730
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x0168f64db4f0e540
    End
    RecordSource ="SELECT top(7) SUM(INVO_LST.MEGHk) AS MEGHT, dbo.Uday(HEAD_LST.DATE_N) AS dd, SUM"
        "(INVO_LST.MABL_K - INVO_LST.N_MOIN) AS MABL_K, HEAD_LST.DATE_N FROM HEAD_LST INN"
        "ER JOIN INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NUMBER AND HEAD_LST.TAG = INVO_LS"
        "T.TAG WHERE (HEAD_LST.TAG = 2) AND (HEAD_LST.TAMIR = - 1) GROUP BY HEAD_LST.DATE"
        "_N, dbo.Uday(HEAD_LST.DATE_N) ORDER BY HEAD_LST.DATE_N DESC"
    Caption ="فروش 7 روزه"
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
        " <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedTyp"
        "e>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>روز</x:Name>\015"
        "\012   <x:SourceName>dd</x:SourceName>\015\012   <x:DataType>Integer</x:DataType"
        ">\015\012   <x:Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adIntege"
        "r</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrde"
        "redMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>مبلغ كل</x:Name>\015\012   <x:SourceName>MABL_K</x:Sour"
        "ceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation"
        "/>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrdered"
        "MembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>روز1</x:Name>\015\012   <x:SourceName>DATE_N<"
        "/x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLCaption>"
        "روز</x:PLCaption>\015\012   <x:FilterCaption>روز</x:FilterCaption>\015\012   <x:"
        "EncodedType>adBigInt</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Unique"
        "Name</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation"
        ">\015\012   <x:Position>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:Pi"
        "votField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of مبلغ كل</x:Name>\015\012"
        "   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012"
        "   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012"
        "   <x:ParentField>مبلغ كل</x:ParentField>\015\012  </x:PivotField>\015\012  <x:P"
        "LTotal>\015\012   <x:Name>Total3</x:Name>\015\012   <x:Caption>Sum of مقدار</x:C"
        "aption>\015\012   <x:ParentField>مقدار</x:ParentField>\015\012   <x:TotalNumber>"
        "1</x:TotalNumber>\015\012  </x:PLTotal>\015\012  <x:PivotData>\015\012   <x:Top>"
        "0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015"
        "\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>504</x:SeqNum>\015\012"
        "  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <"
        "x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012"
        "    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowE"
        "dits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDel"
        "etions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015"
        "\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>9022</"
        "x:Width>\015\012  <x:Height>8493</x:Height>\015\012  <x:DataSource>\015\012   <x"
        ":Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015\012"
        "   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012  <"
        "x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:Cate"
        "gory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "   <x:Data>Total1</x:Data>\015\012  </x:Value>\015\012  <x:BoundCharts>\015\012 "
        "  <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012  <x"
        ":PivotAggOrientation>None</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>\015"
        "\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <"
        "x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entr"
        "y>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000"
        "000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x"
        ":Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry"
        ">\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:En"
        "try>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#6"
        "00080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080<"
        "/x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Ent"
        "ry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015"
        "\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012  "
        " <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:En"
        "try>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:Def"
        "aultFont>\015\012  <x:Border>\015\012   <x:Color>#2F4F4F</x:Color>\015\012  </x:"
        "Border>\015\012  <x:Interior>\015\012   <x:Color>#CCFFCC</x:Color>\015\012   <x:"
        "FillEffect>\015\012    <x:fill x:type=\"Solid\" x:color=\"#CCFFCC\"/>\015\012   "
        "</x:FillEffect>\015\012  </x:Interior>\015\012  <x:Chart>\015\012   <x:PlotArea>"
        "\015\012    <x:Interior>\015\012     <x:Color>#FFFFFF</x:Color>\015\012     <x:F"
        "illEffect>\015\012      <x:fill x:type=\"Solid\" x:color=\"#FFFFFF\"/>\015\012  "
        "   </x:FillEffect>\015\012    </x:Interior>\015\012    <x:Graph>\015\012     <x:"
        "SubType>Smooth</x:SubType>\015\012     <x:Type>Line</x:Type>\015\012     <x:SubT"
        "ype>Standard</x:SubType>\015\012     <x:Series>\015\012      <x:Line>\015\012   "
        "    <x:Color>#000000</x:Color>\015\012      </x:Line>\015\012      <x:FormatMap>"
        "\015\012      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <"
        "x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012      "
        " <x:Data>&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:I"
        "ndex>0</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x"
        ":DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012     "
        "  <x:DataSourceIndex>0</x:DataSourceIndex>\015\012       <x:Data>Total1</x:Data>"
        "\015\012      </x:Value>\015\012      <x:Marker>\015\012       <x:Symbol>None</x"
        ":Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012 "
        "     <x:Thickness>10</x:Thickness>\015\012      <x:Smooth/>\015\012      <x:Data"
        "SourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!</x:Identifier>\015"
        "\012     </x:Series>\015\012     <x:Dimension>\015\012      <x:ScaleID>307459248"
        "</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     </x:Dimensio"
        "n>\015\012     <x:Dimension>\015\012      <x:ScaleID>307467368</x:ScaleID>\015\012"
        "      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimensi"
        "on>\015\012      <x:ScaleID>307467572</x:ScaleID>\015\012      <x:Index>FormatVa"
        "lue</x:Index>\015\012     </x:Dimension>\015\012     <x:Overlap>100</x:Overlap>\015"
        "\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstS"
        "liceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>30724"
        "2060</x:AxisID>\015\012     <x:ScaleID>307459248</x:ScaleID>\015\012     <x:Type"
        ">Category</x:Type>\015\012     <x:Font>\015\012      <x:FontName>Tahoma</x:FontN"
        "ame>\015\012      <x:Size>7</x:Size>\015\012      <x:B>Automatic</x:B>\015\012  "
        "    <x:I>Automatic</x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x:Font>"
        "\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x"
        ":MinorTick>\015\012     <x:Placement>Bottom</x:Placement>\015\012     <x:Groupin"
        "gEnum>Auto</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012    "
        " <x:AxisID>307242944</x:AxisID>\015\012     <x:ScaleID>307467368</x:ScaleID>\015"
        "\012     <x:Type>Value</x:Type>\015\012     <x:Number>\015\012      <x:FormatStr"
        "ing>#,###</x:FormatString>\015\012     </x:Number>\015\012     <x:MajorTick>Outs"
        "ide</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Pla"
        "cement>Left</x:Placement>\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012  "
        " <x:Identifier></x:Identifier>\015\012  </x:Chart>\015\012  <x:Title>\015\012   "
        "<x:Font>\015\012    <x:FontName>B Titr</x:FontName>\015\012    <x:Size>7</x:Size"
        ">\015\012    <x:Color>#A52A2A</x:Color>\015\012    <x:B>Automatic</x:B>\015\012 "
        "   <x:I>Automatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012"
        "   <x:Border>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>"
        "\015\012   <x:Caption>\015\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012"
        "    <x:Data>&quot; فروش 7 روزه&quot;</x:Data>\015\012   </x:Caption>\015\012  </"
        "x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>307459248</x:ScaleID>\015\012"
        "  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>307467368</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>307467572</x:ScaleI"
        "D>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideScreenTi"
        "ps/>\015\012 </x:ChartSpace>\015\012</xml>"
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
                    ColumnWidth =1590
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
                    Top =1474
                    ColumnWidth =390
                    ColumnOrder =0
                    TabIndex =1
                    Name ="DD"
                    ControlSource ="DD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1474
                            Width =480
                            Height =240
                            Name ="Label3"
                            Caption ="روز"
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
                    TabIndex =2
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =1834
                            Width =480
                            Height =240
                            Name ="Label5"
                            Caption ="روز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1133
                    Top =2494
                    ColumnWidth =1590
                    TabIndex =3
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3860
                            Top =2494
                            Width =675
                            Height =240
                            Name ="Label8"
                            Caption ="مبلغ كل"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "amar_frosh7.cls"
