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
    Width =5612
    RowHeight =330
    ItemSuffix =42
    Left =5880
    Top =24675
    Right =7305
    Bottom =26100
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    Tag ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    ShortcutMenuBar ="MNEMUDAR"
    RecSrcDt = Begin
        0x192491f59052e540
    End
    RecordSource ="SELECT SUM(MAND) AS MANDS, col1, ANBAR, SUM(MANDF) AS MANDFS, ANBARN, NAMES, col"
        "2, col3, col4, col5, col6 FROM AKMOGUDI_KOL_ANBAR_N_FANI(99999999, 1) AKMOGUDI_K"
        "OL_ANBAR_N_FANI GROUP BY col1, ANBAR, ANBARN, NAMES, col2, col3, col4, col5, col"
        "6"
    Caption ="موجودي  كالا ها به تفكيك انبار"
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
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>موجودي</x:Name>\015\012   <x:SourceName>MANDS</x:SourceName>\015\012 "
        "  <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedTy"
        "pe>\015\012   <x:NumberFormat>Fixed</x:NumberFormat>\015\012   <x:CompareOrdered"
        "MembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>SAIZE</x:Name>\015\012   <x:EncodedType>adVar"
        "WChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compare"
        "OrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>كد انبار</x:Name>\015\012   <x:SourceName"
        ">ANBAR</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Enc"
        "odedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueNam"
        "e</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:P"
        "ivotField>\015\012  <x:PivotField>\015\012   <x:Name>كارتن</x:Name>\015\012   <x"
        ":SourceName>MANDFS</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:PLDataOrientation/>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:NumberFormat>Fixed</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>انبار</x:Name>\015\012   <x:SourceName>ANBARN</x:Source"
        "Name>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOr"
        "deredMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFi"
        "eldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>نام"
        " واحد</x:Name>\015\012   <x:SourceName>NAMES</x:SourceName>\015\012   <x:Encoded"
        "Type>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName<"
        "/x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>TARH</x:Name>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>"
        "\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>DEGR</x:Name>\015\012   <x:EncodedType>adVarWChar</x:Encoded"
        "Type>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy"
        ">\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>كد</x:Name>\015\012   <x:SourceName>N_FANI</x:SourceNam"
        "e>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrder"
        "edMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>نام كالا</x:Name>\015\012   <x:SourceName>NAM"
        "E</x:SourceName>\015\012   <x:Orientation>Row</x:Orientation>\015\012   <x:Encod"
        "edType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueNam"
        "e</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientation>Co"
        "lumn</x:Orientation>\015\012   <x:Position>-1</x:Position>\015\012   <x:DataFiel"
        "d/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of كا"
        "رتن</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumber>0</x"
        ":TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Positio"
        "n>1</x:Position>\015\012   <x:ParentField>كارتن</x:ParentField>\015\012   <x:Num"
        "berFormat>Fixed</x:NumberFormat>\015\012  </x:PivotField>\015\012  <x:PivotData>"
        "\015\012   <x:Top>0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <"
        "x:Left>0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>5"
        "1</x:SeqNum>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFi"
        "ltered/>\015\012   <x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 1"
        "0.0</x:Caption>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:All"
        "owEdits>false</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:AllowAdditions"
        ">\015\012   <x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:PivotView>\015"
        "\012 </x:PivotTable>\015\012</xml>"
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
        "  <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012     "
        " <x:Name>Series</x:Name>\015\012      <x:Caption>\015\012       <x:DataSourceInd"
        "ex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Series&quot;</x:Data>\015\012"
        "      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012      <x:Category>\015"
        "\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015"
        "\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      </x:Value>\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-"
        "3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatVa"
        "lue>\015\012      <x:DataLabels>\015\012       <x:Border>\015\012        <x:Colo"
        "rIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012       <x:Font>\015\012"
        "        <x:Size>10</x:Size>\015\012        <x:Color>#FF0000</x:Color>\015\012   "
        "     <x:B/>\015\012        <x:I>Automatic</x:I>\015\012        <x:U>Automatic</x"
        ":U>\015\012       </x:Font>\015\012       <x:Number>\015\012        <x:FormatStr"
        "ing>General</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowValu"
        "e/>\015\012       <x:Separator>;</x:Separator>\015\012       <x:Position>Outside"
        "End</x:Position>\015\012      </x:DataLabels>\015\012      <x:Marker>\015\012   "
        "    <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0"
        "</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSour"
        "ceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of كارتن</x:Ident"
        "ifier>\015\012     </x:Series>\015\012     <x:Dimension>\015\012      <x:ScaleID"
        ">372717876</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     </"
        "x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>372718080</x:Scal"
        "eID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012   "
        "  <x:Dimension>\015\012      <x:ScaleID>372718284</x:ScaleID>\015\012      <x:In"
        "dex>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0</x"
        ":Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAngl"
        "e>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     <x"
        ":AxisID>372721472</x:AxisID>\015\012     <x:ScaleID>372717876</x:ScaleID>\015\012"
        "     <x:Type>Category</x:Type>\015\012     <x:Font>\015\012      <x:FontName>Tah"
        "oma</x:FontName>\015\012      <x:B>Automatic</x:B>\015\012      <x:I>Automatic</"
        "x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x:Font>\015\012     <x:Alig"
        "nment>\015\012      <x:Rotation>90</x:Rotation>\015\012     </x:Alignment>\015\012"
        "     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTi"
        "ck>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Tahoma"
        "</x:FontName>\015\012       <x:Size>8</x:Size>\015\012       <x:Color>#000000</x"
        ":Color>\015\012       <x:B>Automatic</x:B>\015\012       <x:I>Automatic</x:I>\015"
        "\012       <x:U>Automatic</x:U>\015\012      </x:Font>\015\012      <x:Border>\015"
        "\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </x:Border>\015\012   "
        "   <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012  "
        "     <x:Data>&quot;نام كالا&quot;</x:Data>\015\012      </x:Caption>\015\012    "
        "  <x:Position>Bottom</x:Position>\015\012     </x:Title>\015\012     <x:Placemen"
        "t>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012"
        "    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>372722868</x:AxisID>\015"
        "\012     <x:ScaleID>372718080</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015"
        "\012     <x:Number>\015\012      <x:FormatString>General</x:FormatString>\015\012"
        "     </x:Number>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridlines>\015"
        "\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:Min"
        "orTick>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Ta"
        "homa</x:FontName>\015\012       <x:Size>8</x:Size>\015\012       <x:B>Automatic<"
        "/x:B>\015\012       <x:I>Automatic</x:I>\015\012       <x:U>Automatic</x:U>\015\012"
        "      </x:Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:Colo"
        "rIndex>\015\012      </x:Border>\015\012      <x:Caption>\015\012       <x:DataS"
        "ourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;موجودي&quot;</x:Da"
        "ta>\015\012      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012"
        "     </x:Title>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>"
        "\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Ch"
        "art>\015\012  <x:Title>\015\012   <x:Font>\015\012    <x:FontName>Sina</x:FontNa"
        "me>\015\012    <x:Color>#228B22</x:Color>\015\012    <x:B>Automatic</x:B>\015\012"
        "    <x:I>Automatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015"
        "\012   <x:Border>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Bor"
        "der>\015\012   <x:Caption>\015\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015"
        "\012    <x:Data>&quot;نمودار موجودي كالا ها&quot;</x:Data>\015\012   </x:Caption"
        ">\015\012  </x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>372717876</x:Sca"
        "leID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>372718080<"
        "/x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>3727"
        "18284</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  "
        "<x:HideScreenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
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
        End
        Begin Section
            CanGrow = NotDefault
            Height =7269
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1749
                    Top =396
                    ColumnWidth =1035
                    ColumnOrder =8
                    Name ="MAND"
                    ControlSource ="MANDS"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4806
                            Top =393
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1746
                    Top =738
                    ColumnWidth =690
                    ColumnOrder =1
                    TabIndex =1
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4503
                            Top =738
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="كد انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1746
                    Top =1098
                    ColumnWidth =1455
                    ColumnOrder =0
                    TabIndex =2
                    Name ="ANBARN"
                    ControlSource ="ANBARN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4503
                            Top =1098
                            Width =645
                            Height =240
                            Name ="Label20"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1755
                    Top =1485
                    ColumnOrder =2
                    TabIndex =3
                    Name ="col1"
                    ControlSource ="col1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4512
                            Top =1485
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
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1755
                    Top =1845
                    ColumnOrder =3
                    TabIndex =4
                    Name ="col2"
                    ControlSource ="col2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4512
                            Top =1845
                            Width =645
                            Height =240
                            Name ="col2l"
                            Caption ="طرح"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1755
                    Top =2130
                    ColumnWidth =510
                    ColumnOrder =4
                    TabIndex =5
                    Name ="col3"
                    ControlSource ="col3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4512
                            Top =2130
                            Width =645
                            Height =240
                            Name ="col3l"
                            Caption ="درجه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =680
                    Top =3354
                    ColumnOrder =9
                    TabIndex =6
                    Name ="MANDF"
                    ControlSource ="MANDFS"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3735
                            Top =3345
                            Width =1035
                            Height =240
                            Name ="Label26"
                            Caption ="كارتن"
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
                    Left =1247
                    Top =2551
                    ColumnWidth =3735
                    ColumnOrder =7
                    TabIndex =7
                    Name ="col22"
                    ControlSource ="col6"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP =6)"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4004
                            Top =2551
                            Width =645
                            Height =240
                            Name ="col5l2"
                            Caption ="طرح"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1360
                    Top =4422
                    ColumnWidth =450
                    ColumnOrder =10
                    TabIndex =8
                    Name ="col4"
                    ControlSource ="col4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4117
                            Top =4422
                            Width =645
                            Height =240
                            Name ="col4l"
                            Caption ="درجه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1360
                    Top =4782
                    ColumnWidth =510
                    ColumnOrder =5
                    TabIndex =9
                    Name ="col5"
                    ControlSource ="col5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4117
                            Top =4782
                            Width =645
                            Height =240
                            Name ="col5l"
                            Caption ="درجه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1360
                    Top =5142
                    ColumnWidth =510
                    ColumnOrder =6
                    TabIndex =10
                    Name ="col6"
                    ControlSource ="col6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4117
                            Top =5142
                            Width =645
                            Height =240
                            Name ="col6l"
                            Caption ="درجه"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =2304
                    Top =5424
                    Width =1701
                    Height =1701
                    TabIndex =11
                    Name ="AK_MOGUDI_ANBAR_LIST_MP7"
                    SourceObject ="Form.AK_MOGUDI_ANBAR_LIST_MP7"
                    LinkChildFields ="col1;col2;col3;col4;col5;col6"
                    LinkMasterFields ="col1;col2;col3;col4;col5;col6"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3345
                            Top =5184
                            Width =660
                            Height =240
                            Name ="Label41"
                            Caption ="Child40:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AK_MOGUDI_ANBAR_LIST_MP6.cls"
