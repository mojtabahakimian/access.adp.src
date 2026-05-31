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
    Left =5610
    Top =660
    Right =10470
    Bottom =5220
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xfc9d76bfd1b8e540
    End
    RecordSource ="nabz_mali30"
    Caption ="مبلغ نقدي"
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
        "   <x:Name>روز</x:Name>\015\012   <x:SourceName>DATE_S</x:SourceName>\015\012   "
        "<x:DataType>Number</x:DataType>\015\012   <x:Orientation>Row</x:Orientation>\015"
        "\012   <x:Position>2</x:Position>\015\012   <x:EncodedType>adBigInt</x:EncodedTy"
        "pe>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   "
        "<x:Name>مبلغ كل</x:Name>\015\012   <x:SourceName>mab</x:SourceName>\015\012   <x"
        ":DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/>\015\012   <x:Enco"
        "dedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName"
        "</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>dd</x:Name>\015\012   <x:DataType>Integer</x:DataType>\015\012   "
        "<x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ماه</x:Name>\015\012"
        "   <x:SourceName>mm</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015"
        "\012   <x:Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adInteger</x:"
        "EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMe"
        "mbersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation"
        ">\015\012   <x:Position>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:Pi"
        "votField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of مبلغ كل</x:Name>\015\012"
        "   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012"
        "   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012"
        "   <x:ParentField>مبلغ كل</x:ParentField>\015\012  </x:PivotField>\015\012  <x:P"
        "ivotData>\015\012   <x:Top>0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015"
        "\012   <x:Left>0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:"
        "SeqNum>214</x:SeqNum>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012   <"
        "x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:Caption>Microsoft Office Piv"
        "otTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012"
        "   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:Allow"
        "Additions>\015\012   <x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:Piv"
        "otView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>14658</x:Height>\015\012  <x:DataSource>\015\012   "
        "<x:Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012"
        "  <x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:C"
        "ategory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012  </x:Value>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:Data"
        "SourceIndex>\015\012  </x:BoundCharts>\015\012  <x:PivotAggOrientation>Series</x"
        ":PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping/>\015"
        "\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:Entry>"
        "\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#8080FF</x:Ent"
        "ry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:Entry>#FFFFA0</x:Entry>\015"
        "\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#600080</x:Entry>\015\012  "
        " <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080</x:Entry>\015\012   <x:En"
        "try>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Entry>\015\012   <x:Entry>#F"
        "F00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015\012   <x:Entry>#0080FF<"
        "/x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#C0FF80</x:Ent"
        "ry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:Entry>#FF80FF</x:Entry>\015"
        "\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:DefaultFont>\015\012  <x:Bor"
        "der>\015\012   <x:Color>#2F4F4F</x:Color>\015\012  </x:Border>\015\012  <x:Inter"
        "ior>\015\012   <x:Color>#FF99CC</x:Color>\015\012   <x:FillEffect>\015\012    <x"
        ":fill x:type=\"Solid\" x:color=\"#FF99CC\"/>\015\012   </x:FillEffect>\015\012  "
        "</x:Interior>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Interior>\015"
        "\012     <x:Color>#FFC0CB</x:Color>\015\012     <x:FillEffect>\015\012      <x:f"
        "ill x:type=\"Solid\" x:color=\"#FFC0CB\"/>\015\012     </x:FillEffect>\015\012  "
        "  </x:Interior>\015\012    <x:Graph>\015\012     <x:SubType>Smooth</x:SubType>\015"
        "\012     <x:Type>Line</x:Type>\015\012     <x:SubType>Standard</x:SubType>\015\012"
        "     <x:Series>\015\012      <x:Line>\015\012       <x:Color>#000000</x:Color>\015"
        "\012      </x:Line>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012"
        "      <x:Name>Series</x:Name>\015\012      <x:Caption>\015\012       <x:DataSour"
        "ceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Series&quot;</x:Data>"
        "\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012      <x:Cat"
        "egory>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:C"
        "ategory>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      </x:Value>\015\012      <x:Marker>\015\012       <x:Symbol>N"
        "one</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015"
        "\012      <x:Thickness>10</x:Thickness>\015\012      <x:Smooth/>\015\012      <x"
        ":DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of مبلغ "
        "كل</x:Identifier>\015\012     </x:Series>\015\012     <x:Dimension>\015\012     "
        " <x:ScaleID>515455008</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015"
        "\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>515455"
        "212</x:ScaleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>"
        "\015\012     <x:Dimension>\015\012      <x:ScaleID>515455416</x:ScaleID>\015\012"
        "      <x:Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:O"
        "verlap>100</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:F"
        "irstSliceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015"
        "\012     <x:AxisID>515458840</x:AxisID>\015\012     <x:ScaleID>515455008</x:Scal"
        "eID>\015\012     <x:Type>Category</x:Type>\015\012     <x:Font>\015\012      <x:"
        "FontName>Tahoma</x:FontName>\015\012      <x:Size>7</x:Size>\015\012      <x:B>A"
        "utomatic</x:B>\015\012      <x:I>Automatic</x:I>\015\012      <x:U>Automatic</x:"
        "U>\015\012     </x:Font>\015\012     <x:Alignment>\015\012      <x:Rotation>90</"
        "x:Rotation>\015\012     </x:Alignment>\015\012     <x:MajorTick>Outside</x:Major"
        "Tick>\015\012     <x:MinorTick>Inside</x:MinorTick>\015\012     <x:Placement>Bot"
        "tom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    <"
        "/x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>515459548</x:AxisID>\015\012 "
        "    <x:ScaleID>515455212</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015\012 "
        "    <x:Font>\015\012      <x:Size>7</x:Size>\015\012      <x:B>Automatic</x:B>\015"
        "\012      <x:I>Automatic</x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x"
        ":Font>\015\012     <x:Number>\015\012      <x:FormatString>#,###</x:FormatString"
        ">\015\012     </x:Number>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012"
        "     <x:MinorTick>None</x:MinorTick>\015\012     <x:Placement>Left</x:Placement>"
        "\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identi"
        "fier>\015\012  </x:Chart>\015\012  <x:Title>\015\012   <x:Font>\015\012    <x:Fo"
        "ntName>B Titr</x:FontName>\015\012    <x:Size>7</x:Size>\015\012    <x:Color>#A5"
        "2A2A</x:Color>\015\012    <x:B>Automatic</x:B>\015\012    <x:I>Automatic</x:I>\015"
        "\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012   <x:Border>\015\012   "
        " <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>\015\012   <x:Caption>\015"
        "\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012    <x:Data>&quot;مبلغ "
        "نقد 30 روزه&quot;</x:Data>\015\012   </x:Caption>\015\012  </x:Title>\015\012  <"
        "x:Scaling>\015\012   <x:ScaleID>515455008</x:ScaleID>\015\012  </x:Scaling>\015\012"
        "  <x:Scaling>\015\012   <x:ScaleID>515455212</x:ScaleID>\015\012  </x:Scaling>\015"
        "\012  <x:Scaling>\015\012   <x:ScaleID>515455416</x:ScaleID>\015\012  </x:Scalin"
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
                    Top =1834
                    ColumnOrder =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
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
                    TabIndex =1
                    Name ="MABL_K"
                    ControlSource ="mab"
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
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    TabIndex =2
                    Name ="mm"
                    ControlSource ="mm"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3027
                            Width =375
                            Height =240
                            Name ="Label10"
                            Caption ="ماه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "nabz_sanad30.cls"
