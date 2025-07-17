function Export-Excel {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([OfficeOpenXml.ExcelPackage])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    param(
        [Parameter(ParameterSetName = 'Default', Position = 0)]
        [String]$Path,
        [Parameter(Mandatory = $true, ParameterSetName = "Package")]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(ValueFromPipeline = $true)]
        [Alias('TargetData')]
        $InputObject,
        [Switch]$Calculate,
        [Switch]$Show,
        [String]$WorksheetName = 'Sheet1',
        [Alias("PW")]
        [String]$Password,
        [switch]$ClearSheet,
        [switch]$Append,
        [String]$Title,
        [OfficeOpenXml.Style.ExcelFillStyle]$TitleFillPattern = 'Solid',
        [Switch]$TitleBold,
        [Int]$TitleSize = 22,
        $TitleBackgroundColor,
        [parameter(DontShow = $true)]
        [Switch]$IncludePivotTable,
        [String]$PivotTableName,
        [String[]]$PivotRows,
        [String[]]$PivotColumns,
        $PivotData,
        [String[]]$PivotFilter,
        [Switch]$PivotDataToColumn,
        [Hashtable]$PivotTableDefinition,
        [Switch]$IncludePivotChart,
        [Alias('ChartType')]
        [OfficeOpenXml.Drawing.Chart.eChartType]$PivotChartType = 'Pie',
        [Switch]$NoLegend,
        [Switch]$ShowCategory,
        [Switch]$ShowPercent,
        [Switch]$AutoSize,
        $MaxAutoSizeRows = 1000,
        [Switch]$NoClobber,
        [Switch]$FreezeTopRow,
        [Switch]$FreezeFirstColumn,
        [Switch]$FreezeTopRowFirstColumn,
        [Int[]]$FreezePane,
        [Switch]$AutoFilter,
        [Switch]$BoldTopRow,
        [Switch]$NoHeader,
        [ValidateScript( {
                if (-not $_) { throw 'RangeName is null or empty.' }
                elseif ($_[0] -notmatch '[a-z]') { throw 'RangeName starts with an invalid character.' }
                else { $true }
            })]
        [String]$RangeName,
        [Alias('Table')]
        $TableName,
        [OfficeOpenXml.Table.TableStyles]$TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium6,
        [HashTable]$TableTotalSettings,
        [Switch]$BarChart,
        [Switch]$PieChart,
        [Switch]$LineChart ,
        [Switch]$ColumnChart ,
        [Object[]]$ExcelChartDefinition,
        [String[]]$HideSheet,
        [String[]]$UnHideSheet,
        [Switch]$MoveToStart,
        [Switch]$MoveToEnd,
        $MoveBefore ,
        $MoveAfter ,
        [Switch]$KillExcel,
        [Switch]$AutoNameRange,
        [Int]$StartRow = 1,
        [Int]$StartColumn = 1,
        [alias('PT')]
        [Switch]$PassThru,
        [String]$Numberformat = 'General',
        [string[]]$ExcludeProperty,
        [Switch]$NoAliasOrScriptPropeties,
        [Switch]$DisplayPropertySet,
        [String[]]$NoNumberConversion,
        [String[]]$NoHyperLinkConversion,
        [Object[]]$ConditionalFormat,
        [Object[]]$ConditionalText,
        [Object[]]$Style,
        [ScriptBlock]$CellStyleSB,
        #If there is already content in the workbook the sheet with the PivotTable will not be active UNLESS Activate is specified
        [switch]$Activate,
        [Parameter(ParameterSetName = 'Default')]
        [Switch]$Now,
        [Switch]$ReturnRange,
        #By default PivotTables have Totals for each Row (on the right) and for each column at the bottom. This allows just one or neither to be selected.
        [ValidateSet("Both", "Columns", "Rows", "None")]
        [String]$PivotTotals = "Both",
        #Included for compatibility - equivalent to -PivotTotals "None"
        [Switch]$NoTotalsInPivot,
        [Switch]$ReZip
    )

    begin {
        $numberRegex = [Regex]'\d'
        $isDataTypeValueType = $false
        if ($NoClobber) { Write-Warning -Message "-NoClobber parameter is no longer used" }
        #Open the file, get the worksheet, and decide where in the sheet we are writing, and if there is a number format to apply.
        try {
            $script:Header = $null
            if ($Append -and $ClearSheet) { throw "You can't use -Append AND -ClearSheet." ; return }
            #To force -Now not to format as a table, allow $false in -TableName to be "No table"
            $TableName = if ($null -eq $TableName -or ($TableName -is [bool] -and $false -eq $TableName)) { $null } else { [String]$TableName }
            if ($Now -or (-not $Path -and -not $ExcelPackage) ) {
                if (-not $PSBoundParameters.ContainsKey("Path")) { $Path = [System.IO.Path]::GetTempFileName() -replace '\.tmp', '.xlsx' }
                if (-not $PSBoundParameters.ContainsKey("Show")) { $Show = $true }
                if (-not $PSBoundParameters.ContainsKey("AutoSize")) { $AutoSize = $true }
                #"Now" option will create a table, unless something passed in TableName/Table Style. False in TableName will block autocreation
                if (-not $PSBoundParameters.ContainsKey("TableName") -and
                    -not $PSBoundParameters.ContainsKey("TableStyle") -and
                    -not $AutoFilter) {
                    $TableName = '' # later rely on distinction between NULL and ""
                }
            }
            if ($ExcelPackage) {
                $pkg = $ExcelPackage
                $Path = $pkg.File
            }
            Else { $pkg = Open-ExcelPackage -Path $Path -Create -KillExcel:$KillExcel -Password:$Password }
        }
        catch { throw "Could not open Excel Package $path" }
        try {
            $params = @{WorksheetName = $WorksheetName }
            foreach ($p in @("ClearSheet", "MoveToStart", "MoveToEnd", "MoveBefore", "MoveAfter", "Activate")) { if ($PSBoundParameters[$p]) { $params[$p] = $PSBoundParameters[$p] } }
            $ws = $pkg | Add-Worksheet @params
            if ($ws.Name -ne $WorksheetName) {
                Write-Warning -Message "The Worksheet name has been changed from $WorksheetName to $($ws.Name), this may cause errors later."
                $WorksheetName = $ws.Name
            }
        }
        catch { throw "Could not get worksheet $WorksheetName" }
        try {
            if ($Append -and $ws.Dimension) {
                #if there is a title or anything else above the header row, append needs to be combined wih a suitable startrow parameter
                $headerRange = $ws.Dimension.Address -replace "\d+$", $StartRow
                #using a slightly odd syntax otherwise header ends up as a 2D array
                $ws.Cells[$headerRange].Value | ForEach-Object -Begin { $Script:header = @() } -Process { $Script:header += $_ }
                $NoHeader = $true
                #if we did not get AutoNameRange, but headers have ranges of the same name make autoNameRange True, otherwise make it false
                if (-not $AutoNameRange) {
                    $AutoNameRange = $true ; foreach ($h in $header) { if ($ws.names.name -notcontains $h) { $AutoNameRange = $false } }
                }
                #if we did not get a Rangename but there is a Range which covers the active part of the sheet, set Rangename to that.
                if (-not $RangeName -and $ws.names.where({ $_.name[0] -match '[a-z]' })) {
                    $theRange = $ws.names.where({
                         ($_.Name[0] -match '[a-z]' ) -and
                         ($_.Start.Row -eq $StartRow) -and
                         ($_.Start.Column -eq $StartColumn) -and
                         ($_.End.Row -eq $ws.Dimension.End.Row) -and
                         ($_.End.Column -eq $ws.Dimension.End.column) } , 'First', 1)
                    if ($theRange) { $rangename = $theRange.name }
                }

                #if we did not get a table name but there is a table which covers the active part of the sheet, set table name to that, and don't do anything with autofilter
                $existingTable = $ws.Tables.Where({ $_.address.address -eq $ws.dimension.address }, 'First', 1)
                if ($null -eq $TableName -and $existingTable) {
                    $TableName = $existingTable.Name
                    $TableStyle = $existingTable.StyleName -replace "^TableStyle", ""
                    $AutoFilter = $false
                }
                #if we did not get $autofilter but a filter range is set and it covers the right area, set autofilter to true
                elseif (-not $AutoFilter -and $ws.Names['_xlnm._FilterDatabase']) {
                    if ( ($ws.Names['_xlnm._FilterDatabase'].Start.Row -eq $StartRow) -and
                         ($ws.Names['_xlnm._FilterDatabase'].Start.Column -eq $StartColumn) -and
                         ($ws.Names['_xlnm._FilterDatabase'].End.Row -eq $ws.Dimension.End.Row) -and
                         ($ws.Names['_xlnm._FilterDatabase'].End.Column -eq $ws.Dimension.End.Column) ) { $AutoFilter = $true }
                }

                $row = $ws.Dimension.End.Row
                Write-Debug -Message ("Appending: headers are " + ($script:Header -join ", ") + " Start row is $row")
                if ($Title) { Write-Warning -Message "-Title Parameter is ignored when appending." }
            }
            elseif ($Title) {
                #Can only add a title if not appending!
                $row = $StartRow
                $ws.Cells[$row, $StartColumn].Value = $Title
                $ws.Cells[$row, $StartColumn].Style.Font.Size = $TitleSize

                if ($PSBoundParameters.ContainsKey("TitleBold")) {
                    #Set title to Bold face font if -TitleBold was specified.
                    #Otherwise the default will be unbolded.
                    $ws.Cells[$row, $StartColumn].Style.Font.Bold = [boolean]$TitleBold
                }
                if ($TitleBackgroundColor ) {
                    if ($TitleBackgroundColor -is [string]) { $TitleBackgroundColor = [System.Drawing.Color]::$TitleBackgroundColor }
                    $ws.Cells[$row, $StartColumn].Style.Fill.PatternType = $TitleFillPattern
                    $ws.Cells[$row, $StartColumn].Style.Fill.BackgroundColor.SetColor($TitleBackgroundColor)
                }
                $row ++ ; $startRow ++
            }
            else { $row = $StartRow }
            $ColumnIndex = $StartColumn
            $Numberformat = Expand-NumberFormat -NumberFormat $Numberformat
            if ((-not $ws.Dimension) -and ($Numberformat -ne $ws.Cells.Style.Numberformat.Format)) {
                $ws.Cells.Style.Numberformat.Format = $Numberformat
                $setNumformat = $false
            }
            else { $setNumformat = ($Numberformat -ne $ws.Cells.Style.Numberformat.Format) }
        }
        catch { throw "Failed preparing to export to worksheet '$WorksheetName' to '$Path': $_" }
        #region Special case -inputobject passed a dataTable object
        <# If inputObject was passed via the pipeline it won't be visible until the process block, we will only see it here if it was passed as a parameter
          if it is a data table don't do foreach on it (slow) - put the whole table in and set dates on date columns,
          set things up for the end block, and skip the process block #>
        if ($InputObject -is [System.Data.DataTable]) {
            if ($Append -and $ws.dimension) {
                $row ++
                $null = $ws.Cells[$row, $StartColumn].LoadFromDataTable($InputObject, $false )
                if ($TableName -or $PSBoundParameters.ContainsKey('TableStyle')) {
                    Add-ExcelTable -Range $ws.Cells[$ws.Dimension] -TableName $TableName -TableStyle $TableStyle -TableTotalSettings $TableTotalSettings
                }
            }
            else {
                #Change TableName if $TableName is non-empty; don't leave caller with a renamed table!
                $orginalTableName = $InputObject.TableName
                if ($PSBoundParameters.ContainsKey("TableName")) {
                    $InputObject.TableName = $TableName
                }
                while ($InputObject.TableName -in $pkg.Workbook.Worksheets.Tables.name) {
                    Write-Warning "Table name $($InputObject.TableName) is not unique, adding '_' to it "
                    $InputObject.TableName += "_"
                }
                #Insert as a table, if Tablestyle didn't arrive as a default, or $TableName non-null - even if empty
                if ($null -ne $TableName -or $PSBoundParameters.ContainsKey("TableStyle")) {
                    $null = $ws.Cells[$row, $StartColumn].LoadFromDataTable($InputObject, (-not $noHeader), $TableStyle )
                    # Workaround for EPPlus not marking the empty row on an empty table as dummy row.
                    if ($InputObject.Rows.Count -eq 0) {
                        ($ws.Tables | Select-Object -Last 1).TableXml.table.SetAttribute('insertRow', 1)
                    }
                }
                else {
                    $null = $ws.Cells[$row, $StartColumn].LoadFromDataTable($InputObject, (-not $noHeader) )
                }
                $InputObject.TableName = $orginalTableName
            }
            foreach ($c in $InputObject.Columns.where({ $_.datatype -eq [datetime] })) {
                Set-ExcelColumn -Worksheet $ws -Column ($c.Ordinal + $StartColumn) -NumberFormat 'Date-Time'
            }
            foreach ($c in $InputObject.Columns.where({ $_.datatype -eq [timespan] })) {
                Set-ExcelColumn -Worksheet $ws -Column ($c.Ordinal + $StartColumn) -NumberFormat '[h]:mm:ss'
            }
            $ColumnIndex += $InputObject.Columns.Count - 1
            if ($noHeader) { $row += $InputObject.Rows.Count - 1 }
            else { $row += $InputObject.Rows.Count }
            $null = $PSBoundParameters.Remove('InputObject')
            $firstTimeThru = $false
        }
        #endregion
        else { $firstTimeThru = $true }
    }

    process {
        if ($PSBoundParameters.ContainsKey("InputObject")) {
            try {
                if ($null -eq $InputObject) { $row += 1 }
                foreach ($TargetData in $InputObject) {
                    if ($firstTimeThru) {
                        $firstTimeThru = $false
                        $isDataTypeValueType = ($null -eq $TargetData) -or ($TargetData.GetType().name -match 'string|timespan|datetime|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort|URI|ExcelHyperLink')
                        if ($isDataTypeValueType ) {
                            $script:Header = @(".")       # dummy value to make sure we go through the "for each name in $header"
                            if (-not $Append) { $row -= 1 } # By default row will be 1, it is incremented before inserting values (so it ends pointing at final row.);  si first data row is 2 - move back up 1 if there is no header .
                        }
                        if ($null -ne $TargetData) { Write-Debug "DataTypeName is '$($TargetData.GetType().name)' isDataTypeValueType '$isDataTypeValueType'" }
                    }
                    #region Add headers - if we are appending, or we have been through here once already we will have the headers
                    if (-not $script:Header) {
                        if ($DisplayPropertySet -and $TargetData.psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames) {
                            $script:Header = $TargetData.psStandardmembers.DefaultDisplayPropertySet.ReferencedPropertyNames.Where( { $_ -notin $ExcludeProperty })
                        }
                        else {
                            if ($NoAliasOrScriptPropeties) { $propType = "Property" } else { $propType = "*" }
                            $script:Header = $TargetData.PSObject.Properties.where( { $_.MemberType -like $propType }).Name
                        }
                        foreach ($exclusion in $ExcludeProperty) { $script:Header = $script:Header -notlike $exclusion }
                        if ($NoHeader) {
                            # Don't push the headers to the spreadsheet
                            $row -= 1
                        }
                        else {
                            $ColumnIndex = $StartColumn
                            foreach ($Name in $script:Header) {
                                $ws.Cells[$row, $ColumnIndex].Value = $Name
                                Write-Verbose "Cell '$row`:$ColumnIndex' add header '$Name'"
                                $ColumnIndex += 1
                            }
                        }
                    }
                    #endregion
                    #region Add non header values
                    $row += 1
                    $ColumnIndex = $StartColumn
                    <#
                 For each item in the header OR for the Data item if this is a simple Type or data table :
                   If it is a date insert with one of Excel's built in formats - recognized as "Date and time to be localized"
                   if it is a timespan insert with a built in format for elapsed hours, minutes and seconds
                   if its  any other numeric insert as is , setting format if need be.
                   Preserve URI, Insert a data table, convert non string objects to string.
                   For strings, check for fomula, URI or Number, before inserting as a string  (ignore nulls) #>
                    foreach ($Name in $script:Header) {
                        if ($isDataTypeValueType) { $v = $TargetData }
                        else { $v = $TargetData.$Name }
                        try {
                            if ($v -is [DateTime]) {
                                $ws.Cells[$row, $ColumnIndex].Value = $v
                                $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = 'm/d/yy h:mm' # This is not a custom format, but a preset recognized as date and localized.
                            }
                            elseif ($v -is [TimeSpan]) {
                                $ws.Cells[$row, $ColumnIndex].Value = $v
                                $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = '[h]:mm:ss'
                            }
                            elseif ($v -is [System.ValueType]) {
                                $ws.Cells[$row, $ColumnIndex].Value = $v
                                if ($setNumformat) { $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                            }
                            elseif ($v -is [uri] -and
                                    $NoHyperLinkConversion -ne '*' -and
                                    $NoHyperLinkConversion -notcontains $Name ) {
                                $ws.Cells[$row, $ColumnIndex].HyperLink = $v
                                $ws.Cells[$row, $ColumnIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                                $ws.Cells[$row, $ColumnIndex].Style.Font.UnderLine = $true
                            }
                            elseif ($v -isnot [String] ) {
                                #Other objects or null.
                                if ($null -ne $v) { $ws.Cells[$row, $ColumnIndex].Value = $v.toString() }
                            }
                            elseif ($v[0] -eq '=') {
                                $ws.Cells[$row, $ColumnIndex].Formula = ($v -replace '^=', '')
                                if ($setNumformat) { $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                            }
                            elseif ( $NoHyperLinkConversion -ne '*' -and # Put the check for 'NoHyperLinkConversion is null' first to skip checking for wellformedstring
                                    $NoHyperLinkConversion -notcontains $Name -and
                                    [System.Uri]::IsWellFormedUriString($v , [System.UriKind]::Absolute)
                                ) { 
                                if ($v -match "^xl://internal/") {
                                    $referenceAddress = $v -replace "^xl://internal/" , ""
                                    $display = $referenceAddress -replace "!A1$"   , ""
                                    $h = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $referenceAddress , $display
                                    $ws.Cells[$row, $ColumnIndex].HyperLink = $h
                                }
                                else { $ws.Cells[$row, $ColumnIndex].HyperLink = $v }   #$ws.Cells[$row, $ColumnIndex].Value = $v.AbsoluteUri
                                $ws.Cells[$row, $ColumnIndex].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                                $ws.Cells[$row, $ColumnIndex].Style.Font.UnderLine = $true
                            }
                            else {
                                $number = $null
                                if ( $NoNumberConversion -ne '*' -and # Check if NoNumberConversion isn't specified. Put this first as it's going to stop the if clause. Quicker than putting regex check first
                                    $numberRegex.IsMatch($v) -and # and if it contains digit(s) - this syntax is quicker than -match for many items and cuts out slow checks for non numbers
                                    $NoNumberConversion -notcontains $Name -and
                                    [Double]::TryParse($v, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number)
                                ) {
                                    $ws.Cells[$row, $ColumnIndex].Value = $number
                                    if ($setNumformat) { $ws.Cells[$row, $ColumnIndex].Style.Numberformat.Format = $Numberformat }
                                }
                                else {
                                    $ws.Cells[$row, $ColumnIndex].Value = $v
                                }
                            }
                        }
                        catch { Write-Warning -Message "Could not insert the '$Name' property at Row $row, Column $ColumnIndex" }
                        $ColumnIndex += 1
                    }
                    $ColumnIndex -= 1 # column index will be the last column whether isDataTypeValueType was true or false
                    #endregion
                }
            }
            catch { throw "Failed exporting data to worksheet '$WorksheetName' to '$Path': $_" }
        }
    }

    end {
        if ($firstTimeThru -and $ws.Dimension) {
            $LastRow = $ws.Dimension.End.Row
            $LastCol = $ws.Dimension.End.Column
            $endAddress = $ws.Dimension.End.Address
        }
        else {
            $LastRow = $row
            $LastCol = $ColumnIndex
            $endAddress = [OfficeOpenXml.ExcelAddress]::GetAddress($LastRow , $LastCol)
        }
        $startAddress = [OfficeOpenXml.ExcelAddress]::GetAddress($StartRow, $StartColumn)
        $dataRange = "{0}:{1}" -f $startAddress, $endAddress

        Write-Debug "Data Range '$dataRange'"
        if ($AutoNameRange) {
            try {
                if (-not $script:header) {
                    # if there aren't any headers, use the the first row of data to name the ranges: this is the last point that headers will be used.
                    $headerRange = $ws.Dimension.Address -replace "\d+$", $StartRow
                    #using a slightly odd syntax otherwise header ends up as a 2D array
                    $ws.Cells[$headerRange].Value | ForEach-Object -Begin { $Script:header = @() } -Process { $Script:header += $_ }
                    if ($PSBoundParameters.ContainsKey($TargetData)) {
                        #if Export was called with data that writes no header start the range at $startRow ($startRow is data)
                        $targetRow = $StartRow
                    }
                    else { $targetRow = $StartRow + 1 }                   #if Export was called without data to add names (assume $startRow is a header) or...
                }                                                         #          ... called with data that writes a header, then start the range at $startRow + 1
                else { $targetRow = $StartRow + 1 }

                #Dimension.start.row always seems to be one so we work out the target row
                #, but start.column is the first populated one and .Columns is the count of populated ones.
                # if we have 5 columns from 3 to 8, headers are numbered 0..4, so that is in the for loop and used for getting the name...
                # but we have to add the start column on when referencing positions
                foreach ($c in 0..($LastCol - $StartColumn)) {
                    $targetRangeName = @($script:Header)[$c]  #Let Add-ExcelName fix (and warn about) bad names
                    Add-ExcelName  -RangeName $targetRangeName -Range $ws.Cells[$targetRow, ($StartColumn + $c ), $LastRow, ($StartColumn + $c )]
                    try {
                        #this test can throw with some names, surpress any error
                        if ([OfficeOpenXml.FormulaParsing.ExcelUtilities.ExcelAddressUtil]::IsValidAddress(($targetRangeName -replace '\W' , '_' ))) {
                            Write-Warning -Message "AutoNameRange: Property name '$targetRangeName' is also a valid Excel address and may cause issues. Consider renaming the property."
                        }
                    }
                    catch {
                        Write-Warning -Message "AutoNameRange: Testing '$targetRangeName' caused an error. This should be harmless, but a change of property name may be needed.."
                    }
                }
            }
            catch { Write-Warning -Message "Failed adding named ranges to worksheet '$WorksheetName': $_" }
        }
        #Empty string is not allowed as a name for ranges or tables.
        if ($RangeName) { Add-ExcelName  -Range $ws.Cells[$dataRange] -RangeName $RangeName }

        #Allow table to be inserted by specifying Name, or Style or both; only process autoFilter if there is no table (they clash).
        if ($null -ne $TableName -or $PSBoundParameters.ContainsKey('TableStyle')) {
            #Already inserted Excel table if input was a DataTable
            if ($InputObject -isnot [System.Data.DataTable]) {
                Add-ExcelTable -Range $ws.Cells[$dataRange] -TableName $TableName -TableStyle $TableStyle -TableTotalSettings $TableTotalSettings
            }
        }
        elseif ($AutoFilter) {
            try {
                $ws.Cells[$dataRange].AutoFilter = $true
                Write-Verbose -Message "Enabled autofilter. "
            }
            catch { Write-Warning -Message "Failed adding autofilter to worksheet '$WorksheetName': $_" }
        }

        if ($PivotTableDefinition) {
            foreach ($item in $PivotTableDefinition.GetEnumerator()) {
                $params = $item.value
                if ($Activate) { $params.Activate = $true }
                if ($params.keys -notcontains 'SourceRange' -and
                   ($params.Keys -notcontains 'SourceWorksheet' -or $params.SourceWorksheet -eq $WorksheetName)) { $params.SourceRange = $dataRange }
                if ($params.Keys -notcontains 'SourceWorksheet') { $params.SourceWorksheet = $ws }
                if ($params.Keys -notcontains 'NoTotalsInPivot' -and $NoTotalsInPivot  ) { $params.PivotTotals = 'None' }
                if ($params.Keys -notcontains 'PivotTotals' -and $PivotTotals      ) { $params.PivotTotals = $PivotTotals }
                if ($params.Keys -notcontains 'PivotDataToColumn' -and $PivotDataToColumn) { $params.PivotDataToColumn = $true }

                Add-PivotTable -ExcelPackage $pkg -PivotTableName $item.key @Params
            }
        }
        if ($IncludePivotTable -or $IncludePivotChart -or $PivotData) {
            $params = @{
                'SourceRange' = $dataRange
            }
            if ($PivotTableName -and ($pkg.workbook.worksheets.tables.name -contains $PivotTableName)) {
                Write-Warning -Message "The selected PivotTable name '$PivotTableName' is already used as a table name. Adding a suffix of 'Pivot'."
                $PivotTableName += 'Pivot'
            }

            if ($PivotTableName) { $params.PivotTableName = $PivotTableName }
            else { $params.PivotTableName = $WorksheetName + 'PivotTable' }
            if ($Activate) { $params.Activate = $true }
            if ($PivotFilter) { $params.PivotFilter = $PivotFilter }
            if ($PivotRows) { $params.PivotRows = $PivotRows }
            if ($PivotColumns) { $Params.PivotColumns = $PivotColumns }
            if ($PivotData) { $Params.PivotData = $PivotData }
            if ($NoTotalsInPivot) { $params.PivotTotals = "None" }
            Elseif ($PivotTotals) { $params.PivotTotals = $PivotTotals }
            if ($PivotDataToColumn) { $params.PivotDataToColumn = $true }
            if ($IncludePivotChart -or
                $PSBoundParameters.ContainsKey('PivotChartType')) {
                $params.IncludePivotChart = $true
                $Params.ChartType = $PivotChartType
                if ($ShowCategory) { $params.ShowCategory = $true }
                if ($ShowPercent) { $params.ShowPercent = $true }
                if ($NoLegend) { $params.NoLegend = $true }
            }
            Add-PivotTable -ExcelPackage $pkg -SourceWorksheet $ws   @params
        }

        try {
            #Allow single switch or two seperate ones.
            if ($FreezeTopRowFirstColumn -or ($FreezeTopRow -and $FreezeFirstColumn)) {
                if ($Title) {
                    $ws.View.FreezePanes(3, 2)
                    Write-Verbose -Message "Froze title and header rows and first column"
                }
                else {
                    $ws.View.FreezePanes(2, 2)
                    Write-Verbose -Message "Froze top row and first column"
                }
            }
            elseif ($FreezeTopRow) {
                if ($Title) {
                    $ws.View.FreezePanes(3, 1)
                    Write-Verbose -Message "Froze title and header rows"
                }
                else {
                    $ws.View.FreezePanes(2, 1)
                    Write-Verbose -Message "Froze top row"
                }
            }
            elseif ($FreezeFirstColumn) {
                $ws.View.FreezePanes(1, 2)
                Write-Verbose -Message "Froze first column"
            }
            #Must be 1..maxrows or and array of 1..maxRows,1..MaxCols
            if ($FreezePane) {
                $freezeRow, $freezeColumn = $FreezePane
                if (-not $freezeColumn -or $freezeColumn -eq 0) {
                    $freezeColumn = 1
                }

                if ($freezeRow -ge 1) {
                    $ws.View.FreezePanes($freezeRow, $freezeColumn)
                    Write-Verbose -Message "Froze panes at row $freezeRow and column $FreezeColumn"
                }
            }
        }
        catch { Write-Warning -Message "Failed adding Freezing the panes in worksheet '$WorksheetName': $_" }

        if ($PSBoundParameters.ContainsKey("BoldTopRow")) {
            #it sets bold as far as there are populated cells: for whole row could do $ws.row($x).style.font.bold = $true
            try {
                if ($Title) {
                    $range = $ws.Dimension.Address -replace '\d+', ($StartRow + 1)
                }
                else {
                    $range = $ws.Dimension.Address -replace '\d+', $StartRow
                }
                $ws.Cells[$range].Style.Font.Bold = [boolean]$BoldTopRow
                Write-Verbose -Message "Set $range font style to bold."
            }
            catch { Write-Warning -Message "Failed setting the top row to bold in worksheet '$WorksheetName': $_" }
        }
        if ($AutoSize -and -not $env:NoAutoSize) {
            try {
                #Don't fit the all the columns in the sheet; if we are adding cells beside things with hidden columns, that unhides them
                if ($MaxAutoSizeRows -and $MaxAutoSizeRows -lt $LastRow ) {
                    $AutosizeRange = [OfficeOpenXml.ExcelAddress]::GetAddress($startRow, $StartColumn, $MaxAutoSizeRows , $LastCol)
                    $ws.Cells[$AutosizeRange].AutoFitColumns()
                }
                else { $ws.Cells[$dataRange].AutoFitColumns() }
                Write-Verbose -Message "Auto-sized columns"
            }
            catch { Write-Warning -Message "Failed autosizing columns of worksheet '$WorksheetName': $_" }
        }
        elseif ($AutoSize) { Write-Warning -Message "Auto-fitting columns is not available with this OS configuration." }

        foreach ($Sheet in $HideSheet) {
            try {
                $pkg.Workbook.Worksheets.Where({ $_.Name -like $sheet }) | ForEach-Object {
                    $_.Hidden = 'Hidden'
                    Write-verbose -Message "Sheet '$($_.Name)' Hidden."
                }
            }
            catch { Write-Warning -Message  "Failed hiding worksheet '$sheet': $_" }
        }
        foreach ($Sheet in $UnHideSheet) {
            try {
                $pkg.Workbook.Worksheets.Where({ $_.Name -like $sheet }) | ForEach-Object {
                    $_.Hidden = 'Visible'
                    Write-verbose -Message "Sheet '$($_.Name)' shown"
                }
            }
            catch { Write-Warning -Message  "Failed showing worksheet '$sheet': $_" }
        }
        if (-not $pkg.Workbook.Worksheets.Where({ $_.Hidden -eq 'visible' })) {
            Write-Verbose -Message "No Sheets were left visible, making $WorksheetName visible"
            $ws.Hidden = 'Visible'
        }

        foreach ($chartDef in $ExcelChartDefinition) {
            if ($chartDef -is [System.Management.Automation.PSCustomObject]) {
                $params = @{}
                $chartDef.PSObject.Properties | ForEach-Object { if ( $null -ne $_.value) { $params[$_.name] = $_.value } }
                Add-ExcelChart -Worksheet $ws @params
            }
            elseif ($chartDef -is [hashtable] -or $chartDef -is [System.Collections.Specialized.OrderedDictionary]) {
                Add-ExcelChart -Worksheet $ws @chartDef
            }
        }

        if ($Calculate) {
            try { [OfficeOpenXml.CalculationExtension]::Calculate($ws) }
            catch { Write-Warning "One or more errors occured while calculating, save will continue, but there may be errors in the workbook. $_" }
        }

        if ($Barchart -or $PieChart -or $LineChart -or $ColumnChart) {
            if ($NoHeader) { $FirstDataRow = $startRow }
            else { $FirstDataRow = $startRow + 1 }
            $range = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $startColumn, $FirstDataRow, $lastCol )
            $xCol = $ws.cells[$range] | Where-Object { $_.value -is [string] } | ForEach-Object { $_.start.column } | Sort-Object | Select-Object -first 1
            if (-not $xcol) {
                $xcol = $StartColumn
                $range = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, ($startColumn + 1), $FirstDataRow, $lastCol )
            }
            $yCol = $ws.cells[$range] | Where-Object { $_.value -is [valueType] -or $_.Formula } | ForEach-Object { $_.start.column } | Sort-Object | Select-Object -first 1
            if (-not ($xCol -and $ycol)) { Write-Warning -Message "Can't identify a string column and a number column to use as chart labels and data. " }
            else {
                $params = @{
                    XRange = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $xcol , $lastrow, $xcol)
                    YRange = [OfficeOpenXml.ExcelAddress]::GetAddress($FirstDataRow, $ycol , $lastrow, $ycol)
                    Title  = ''
                    Column = ($lastCol + 1)
                    Width  = 800
                }
                if ($ShowPercent) { $params["ShowPercent"] = $true }
                if ($ShowCategory) { $params["ShowCategory"] = $true }
                if ($NoLegend) { $params["NoLegend"] = $true }
                if (-not $NoHeader) { $params["SeriesHeader"] = $ws.Cells[$startRow, $YCol].Value }
                if ($ColumnChart) { $Params["chartType"] = "ColumnStacked" }
                elseif ($Barchart) { $Params["chartType"] = "BarStacked" }
                elseif ($PieChart) { $Params["chartType"] = "PieExploded3D" }
                elseif ($LineChart) { $Params["chartType"] = "Line" }

                Add-ExcelChart -Worksheet $ws @params
            }
        }

        # It now doesn't matter if the conditional formating rules are passed in $conditionalText or $conditional format.
        # Just one with an alias for compatiblity it will break things for people who are using both at once
        foreach ($c in  (@() + $ConditionalText + $ConditionalFormat) ) {
            try {
                #we can take an object with a .ConditionalType property made by New-ConditionalText or with a .Formatter Property made by New-ConditionalFormattingIconSet or a hash table
                if ($c.ConditionalType) {
                    $cfParams = @{RuleType = $c.ConditionalType; ConditionValue = $c.Text ;
                        BackgroundColor = $c.BackgroundColor; BackgroundPattern = $c.PatternType  ;
                        ForeGroundColor = $c.ConditionalTextColor
                    }
                    if ($c.Range) { $cfParams.Range = $c.Range }
                    else { $cfParams.Range = $ws.Dimension.Address }
                    Add-ConditionalFormatting -Worksheet $ws @cfParams
                    Write-Verbose -Message "Added conditional formatting to range $($c.range)"
                }
                elseif ($c.formatter) {
                    switch ($c.formatter) {
                        "ThreeIconSet" { Add-ConditionalFormatting -Worksheet $ws -ThreeIconsSet $c.IconType -range $c.range -reverse:$c.reverse -ShowIconOnly:$c.ShowIconOnly}
                        "FourIconSet" { Add-ConditionalFormatting -Worksheet $ws  -FourIconsSet $c.IconType -range $c.range -reverse:$c.reverse -ShowIconOnly:$c.ShowIconOnly}
                        "FiveIconSet" { Add-ConditionalFormatting -Worksheet $ws  -FiveIconsSet $c.IconType -range $c.range -reverse:$c.reverse -ShowIconOnly:$c.ShowIconOnly}
                    }
                    Write-Verbose -Message "Added conditional formatting to range $($c.range)"
                }
                elseif ($c -is [hashtable] -or $c -is [System.Collections.Specialized.OrderedDictionary]) {
                    if (-not $c.Range -or $c.Address) { $c.Address = $ws.Dimension.Address }
                    Add-ConditionalFormatting -Worksheet $ws @c
                }
            }
            catch { throw "Error applying conditional formatting to worksheet $_" }
        }
        foreach ($s in $Style) {
            if (-not $s.Range) { $s["Range"] = $ws.Dimension.Address }
            Set-ExcelRange -Worksheet $ws @s
        }
        if ($CellStyleSB) {
            try {
                $TotalRows = $ws.Dimension.Rows
                $LastColumn = $ws.Dimension.Address -replace "^.*:(\w*)\d+$" , '$1'
                & $CellStyleSB $ws $TotalRows $LastColumn
            }
            catch { Write-Warning -Message "Failed processing CellStyleSB in worksheet '$WorksheetName': $_" }
        }

        #Can only add password, may want to support -password $Null removing password.
        if ($Password) {
            try {
                $ws.Protection.SetPassword($Password)
                Write-Verbose -Message 'Set password on workbook'
            }
            catch { throw "Failed setting password for worksheet '$WorksheetName': $_" }
        }

        if ($PassThru) { $pkg }
        else {
            if ($ReturnRange) { $dataRange }

            if ($Password) { $pkg.Save($Password) }
            else { $pkg.Save() }
            Write-Verbose -Message "Saved workbook $($pkg.File)"
            if ($ReZip) {
                Invoke-ExcelReZipFile -ExcelPackage $pkg
            }

            $pkg.Dispose()

            if ($Show) { Invoke-Item $Path }
        }
    }
}

# SIG # Begin signature block
# MIIuzQYJKoZIhvcNAQcCoIIuvjCCLroCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAdAS7KxlGPDdgV
# mgZRCDsGh+zCSTO81GAWe2cj6ab5m6CCE24wggVyMIIDWqADAgECAhB2U/6sdUZI
# k/Xl10pIOk74MA0GCSqGSIb3DQEBDAUAMFMxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENvZGUgU2ln
# bmluZyBSb290IFI0NTAeFw0yMDAzMTgwMDAwMDBaFw00NTAzMTgwMDAwMDBaMFMx
# CzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQD
# EyBHbG9iYWxTaWduIENvZGUgU2lnbmluZyBSb290IFI0NTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALYtxTDdeuirkD0DcrA6S5kWYbLl/6VnHTcc5X7s
# k4OqhPWjQ5uYRYq4Y1ddmwCIBCXp+GiSS4LYS8lKA/Oof2qPimEnvaFE0P31PyLC
# o0+RjbMFsiiCkV37WYgFC5cGwpj4LKczJO5QOkHM8KCwex1N0qhYOJbp3/kbkbuL
# ECzSx0Mdogl0oYCve+YzCgxZa4689Ktal3t/rlX7hPCA/oRM1+K6vcR1oW+9YRB0
# RLKYB+J0q/9o3GwmPukf5eAEh60w0wyNA3xVuBZwXCR4ICXrZ2eIq7pONJhrcBHe
# OMrUvqHAnOHfHgIB2DvhZ0OEts/8dLcvhKO/ugk3PWdssUVcGWGrQYP1rB3rdw1G
# R3POv72Vle2dK4gQ/vpY6KdX4bPPqFrpByWbEsSegHI9k9yMlN87ROYmgPzSwwPw
# jAzSRdYu54+YnuYE7kJuZ35CFnFi5wT5YMZkobacgSFOK8ZtaJSGxpl0c2cxepHy
# 1Ix5bnymu35Gb03FhRIrz5oiRAiohTfOB2FXBhcSJMDEMXOhmDVXR34QOkXZLaRR
# kJipoAc3xGUaqhxrFnf3p5fsPxkwmW8x++pAsufSxPrJ0PBQdnRZ+o1tFzK++Ol+
# A/Tnh3Wa1EqRLIUDEwIrQoDyiWo2z8hMoM6e+MuNrRan097VmxinxpI68YJj8S4O
# JGTfAgMBAAGjQjBAMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBQfAL9GgAr8eDm3pbRD2VZQu86WOzANBgkqhkiG9w0BAQwFAAOCAgEA
# Xiu6dJc0RF92SChAhJPuAW7pobPWgCXme+S8CZE9D/x2rdfUMCC7j2DQkdYc8pzv
# eBorlDICwSSWUlIC0PPR/PKbOW6Z4R+OQ0F9mh5byV2ahPwm5ofzdHImraQb2T07
# alKgPAkeLx57szO0Rcf3rLGvk2Ctdq64shV464Nq6//bRqsk5e4C+pAfWcAvXda3
# XaRcELdyU/hBTsz6eBolSsr+hWJDYcO0N6qB0vTWOg+9jVl+MEfeK2vnIVAzX9Rn
# m9S4Z588J5kD/4VDjnMSyiDN6GHVsWbcF9Y5bQ/bzyM3oYKJThxrP9agzaoHnT5C
# JqrXDO76R78aUn7RdYHTyYpiF21PiKAhoCY+r23ZYjAf6Zgorm6N1Y5McmaTgI0q
# 41XHYGeQQlZcIlEPs9xOOe5N3dkdeBBUO27Ql28DtR6yI3PGErKaZND8lYUkqP/f
# obDckUCu3wkzq7ndkrfxzJF0O2nrZ5cbkL/nx6BvcbtXv7ePWu16QGoWzYCELS/h
# AtQklEOzFfwMKxv9cW/8y7x1Fzpeg9LJsy8b1ZyNf1T+fn7kVqOHp53hWVKUQY9t
# W76GlZr/GnbdQNJRSnC0HzNjI3c/7CceWeQIh+00gkoPP/6gHcH1Z3NFhnj0qinp
# J4fGGdvGExTDOUmHTaCX4GUT9Z13Vunas1jHOvLAzYIwggbmMIIEzqADAgECAhB3
# vQ4DobcI+FSrBnIQ2QRHMA0GCSqGSIb3DQEBCwUAMFMxCzAJBgNVBAYTAkJFMRkw
# FwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENv
# ZGUgU2lnbmluZyBSb290IFI0NTAeFw0yMDA3MjgwMDAwMDBaFw0zMDA3MjgwMDAw
# MDBaMFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8w
# LQYDVQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANZCTfnjT8Yj9GwdgaYw90g9
# z9DljeUgIpYHRDVdBs8PHXBg5iZU+lMjYAKoXwIC947Jbj2peAW9jvVPGSSZfM8R
# Fpsfe2vSo3toZXer2LEsP9NyBjJcW6xQZywlTVYGNvzBYkx9fYYWlZpdVLpQ0LB/
# okQZ6dZubD4Twp8R1F80W1FoMWMK+FvQ3rpZXzGviWg4QD4I6FNnTmO2IY7v3Y2F
# QVWeHLw33JWgxHGnHxulSW4KIFl+iaNYFZcAJWnf3sJqUGVOU/troZ8YHooOX1Re
# veBbz/IMBNLeCKEQJvey83ouwo6WwT/Opdr0WSiMN2WhMZYLjqR2dxVJhGaCJedD
# CndSsZlRQv+hst2c0twY2cGGqUAdQZdihryo/6LHYxcG/WZ6NpQBIIl4H5D0e6lS
# TmpPVAYqgK+ex1BC+mUK4wH0sW6sDqjjgRmoOMieAyiGpHSnR5V+cloqexVqHMRp
# 5rC+QBmZy9J9VU4inBDgoVvDsy56i8Te8UsfjCh5MEV/bBO2PSz/LUqKKuwoDy3K
# 1JyYikptWjYsL9+6y+JBSgh3GIitNWGUEvOkcuvuNp6nUSeRPPeiGsz8h+WX4VGH
# aekizIPAtw9FbAfhQ0/UjErOz2OxtaQQevkNDCiwazT+IWgnb+z4+iaEW3VCzYkm
# eVmda6tjcWKQJQ0IIPH/AgMBAAGjggGuMIIBqjAOBgNVHQ8BAf8EBAMCAYYwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU
# 2rONwCSQo2t30wygWd0hZ2R2C3gwHwYDVR0jBBgwFoAUHwC/RoAK/Hg5t6W0Q9lW
# ULvOljswgZMGCCsGAQUFBwEBBIGGMIGDMDkGCCsGAQUFBzABhi1odHRwOi8vb2Nz
# cC5nbG9iYWxzaWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUwRgYIKwYBBQUHMAKG
# Omh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2NvZGVzaWduaW5n
# cm9vdHI0NS5jcnQwQQYDVR0fBDowODA2oDSgMoYwaHR0cDovL2NybC5nbG9iYWxz
# aWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3JsMFYGA1UdIARPME0wQQYJKwYB
# BAGgMgEyMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29t
# L3JlcG9zaXRvcnkvMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEACIhyJsav
# +qxfBsCqjJDa0LLAopf/bhMyFlT9PvQwEZ+PmPmbUt3yohbu2XiVppp8YbgEtfjr
# y/RhETP2ZSW3EUKL2Glux/+VtIFDqX6uv4LWTcwRo4NxahBeGQWn52x/VvSoXMNO
# Ca1Za7j5fqUuuPzeDsKg+7AE1BMbxyepuaotMTvPRkyd60zsvC6c8YejfzhpX0FA
# Z/ZTfepB7449+6nUEThG3zzr9s0ivRPN8OHm5TOgvjzkeNUbzCDyMHOwIhz2hNab
# XAAC4ShSS/8SS0Dq7rAaBgaehObn8NuERvtz2StCtslXNMcWwKbrIbmqDvf+28rr
# vBfLuGfr4z5P26mUhmRVyQkKwNkEcUoRS1pkw7x4eK1MRyZlB5nVzTZgoTNTs/Z7
# KtWJQDxxpav4mVn945uSS90FvQsMeAYrz1PYvRKaWyeGhT+RvuB4gHNU36cdZytq
# tq5NiYAkCFJwUPMB/0SuL5rg4UkI4eFb1zjRngqKnZQnm8qjudviNmrjb7lYYuA2
# eDYB+sGniXomU6Ncu9Ky64rLYwgv/h7zViniNZvY/+mlvW1LWSyJLC9Su7UpkNpD
# R7xy3bzZv4DB3LCrtEsdWDY3ZOub4YUXmimi/eYI0pL/oPh84emn0TCOXyZQK8ei
# 4pd3iu/YTT4m65lAYPM8Zwy2CHIpNVOBNNwwggcKMIIE8qADAgECAgxwaYJwgzTu
# /lOAy3UwDQYJKoZIhvcNAQELBQAwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEds
# b2JhbFNpZ24gbnYtc2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2Rl
# U2lnbmluZyBDQSAyMDIwMB4XDTI1MDEyOTE1NDU0MloXDTI2MDMwMTE1NDU0Mlow
# eDELMAkGA1UEBhMCVVMxDzANBgNVBAgTBk9yZWdvbjEUMBIGA1UEBxMLTGFrZSBP
# c3dlZ28xIDAeBgNVBAoTF01pbGVzdG9uZSBTeXN0ZW1zLCBJbmMuMSAwHgYDVQQD
# ExdNaWxlc3RvbmUgU3lzdGVtcywgSW5jLjCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBAIfjLyk2i3fyML6J7+LzzwNUZc9/22fhwSDPtkRfLjy8yYgTxbSu
# TgJi/twyUkts2qjbnGqD/maukegK3uZZHHOV2dWL1fjTOa4Mc0Vw3jBELmvrW/CO
# TSvPQMS3u2WVGGJBqVu+IVwWkC2dc2qnNeN4K1VXr4TqCjE0st/K2cJdRbEavMpz
# BCqYdfA41feH7s22C3C0JE5aXVuCaeXbJB0uBfH4fcLn/Wt3wQuT6sMsjSpFj5gH
# iwt4PPP4jWt12kB1LzOzMVUMNfIXrfECqEdRN7lfyZPWTsEiiamjXZEmsxH7Q0rn
# QVvfeh12vf5SrUZFKyi5Gj98AkrWJIciDohd3zRxB/tdSTKF8cVkYwT80/6IOyK1
# y8WaBEO6myXNixx8VDZp20teUSxQex0WgTNQ0/raNWT21ZbKe1SHg/zziHaize6K
# ZM/TikiFz9ibppZR87/lxWw8Oy8eMckjWNDJYitpELcJ78PkBQGISOtGDbH6JbNW
# 2gJ/idTb6hasv/Dvu9QIbyOu+gTm3pJ2Ot5E0nlLGEqem3F2VBoOPLfgsmXfw0S0
# 8IiYGm9QTPy5T15rJNhQHnX0Q7WgPLaictxsOvO5vHXlKln6MRv4dYSYAWUx5Kgy
# 6cptfVXqH1MSf66781BjX3W7a5ySps2O8eWzd/H80tsNF64drkn4kAfdAgMBAAGj
# ggGxMIIBrTAOBgNVHQ8BAf8EBAMCB4AwgZsGCCsGAQUFBwEBBIGOMIGLMEoGCCsG
# AQUFBzAChj5odHRwOi8vc2VjdXJlLmdsb2JhbHNpZ24uY29tL2NhY2VydC9nc2dj
# Y3I0NWNvZGVzaWduY2EyMDIwLmNydDA9BggrBgEFBQcwAYYxaHR0cDovL29jc3Au
# Z2xvYmFsc2lnbi5jb20vZ3NnY2NyNDVjb2Rlc2lnbmNhMjAyMDBWBgNVHSAETzBN
# MEEGCSsGAQQBoDIBMjA0MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxz
# aWduLmNvbS9yZXBvc2l0b3J5LzAIBgZngQwBBAEwCQYDVR0TBAIwADBFBgNVHR8E
# PjA8MDqgOKA2hjRodHRwOi8vY3JsLmdsb2JhbHNpZ24uY29tL2dzZ2NjcjQ1Y29k
# ZXNpZ25jYTIwMjAuY3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB8GA1UdIwQYMBaA
# FNqzjcAkkKNrd9MMoFndIWdkdgt4MB0GA1UdDgQWBBQOXMn2R0Y1ZfbEfoeRqdTg
# piHkQjANBgkqhkiG9w0BAQsFAAOCAgEAat2qjMH8QzC3NJg+KU5VmeEymyAxDaqX
# VUrlYjMR9JzvTqcKrXvhTO8SwM37TW4UR7kmG1TKdVkKvbBAThjEPY5xH+TMgxSQ
# TrOTmBPx2N2bMUoRqNCuf4RmJXGk2Qma26aU/gSKqC2hq0+fstuRQjINUVTp00VS
# 9XlAK0zcXFrZrVREaAr9p4U606oYWT5oen7Bi1M7L8hbbZojw9N4WwuH6n0Ctwlw
# ZUHDbYzXWnKALkPWNHZcqZX4zHFDGKzn/wgu8VDkhaPrmn9lRW+BDyI/EE0iClGJ
# KXMidK71y7BT+DLGBvga6xfbInj5e+n7Nu9gU8D9RkjQqdoq7mO/sXUIsHMJQPkD
# qnyy20PsMYlEbuRTJ8v9c7HN+oWlttEq4iJ24EerC/1Lamr55L/F9rcJ3XRYLKLR
# E5K9pOsq7oJpbsOuYfpv61WISX3ewy5v1tY9VHjn7NQxoMgczSuAf97qbQNpblt0
# h+9KTiwWmLnw1jP1/vwNBYwZk3mtL0Z7Z/l2qqVawrT2W3/EwovP0DWcQr9idTAI
# WLbnWRUHlUZv4rCeoIwXWGgCUOF+BHU1sacps1V3kK1OpNvRWYs+mk2tzGyxoIEB
# 9whM6Vxzik4Q7ciXG7G2ZzYR9f2J4kbr4hTxZEB6ysCPT8DTdqmUwUcxDc2i2j2p
# AkA8F4fhc3Yxghq1MIIasQIBATBpMFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBH
# bG9iYWxTaWduIG52LXNhMS8wLQYDVQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29k
# ZVNpZ25pbmcgQ0EgMjAyMAIMcGmCcIM07v5TgMt1MA0GCWCGSAFlAwQCAQUAoIGk
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCConIL0wQswZMJY5PmaD7tpYCqX+LgV
# q33GZLQBSlBr8DA4BgorBgEEAYI3AgEMMSowKKACgAChIoAgaHR0cHM6Ly93d3cu
# bWlsZXN0b25lcHN0b29scy5jb20wDQYJKoZIhvcNAQEBBQAEggIAH7zwUumhw/B7
# /N1FWcoQC+l2Z9nzqSLkPj3BDFhBtqQs4VdiHzzbL+aSqmsPXUYMXthWUMP9jGzv
# 8SeQjmZyPS/pbK2eyESZyVAuIOsAiK4WrIw6zEW5vYvkdQ3A4x96EuQBohG/Q2Bq
# 449ncanZ0xNntp0wGmr00p9cu4XnAL0J0g/7HqRARLVzqqYkDQPFCITIZRD45SBY
# fvlHYdTthksTdO/xY+3L9co7cAt0aHqMrOY2IZTNofhoYL1N8AVP4GpmOiizeOA+
# Yr9bUqOSQ9kqA667xas+MFImw36TiX/gHg/EPxooUMUObNcPcdJe3m8NS+6QumS0
# M6/+hffwK445J8HaZp2Mjz78GYDFPSMyqhHG+h6kcbX3y0T+oqLATEYtrxtpU/3T
# YAJFKLCBtCEysnH1W0cMAAEurj7/BiannukWQS57xHVrmCtN7E48cQILUtMzjqQA
# t5JySsAcMOlJqyYXdWs32A1U9PpNWRILTzMtUbfQrJINh2Q06WUaRfsS5gNAL9g1
# Os012gO2ggYP8WJ7nqDE+d7LySGnXL0NBdfqgjUjp6E4VHYBZUe/5yMdBJH8pcSY
# sAJeUs0vPiiMGx7V2+XyIEsKVqO1uOlw8zbmxnjlNuCLNwggcYHHkrlYww2Dt1f1
# 7SbdC73ZI7nHGjCMjD1kOO26cdwVVHChghd2MIIXcgYKKwYBBAGCNwMDATGCF2Iw
# ghdeBgkqhkiG9w0BBwKgghdPMIIXSwIBAzEPMA0GCWCGSAFlAwQCAQUAMHcGCyqG
# SIb3DQEJEAEEoGgEZjBkAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# iGWWIS+xetWBx2cl09uu76S0xldo2t4GtnxosMYSPqsCEH7u1StHXEt6my9yEKaZ
# oIIYDzIwMjUwNjE4MTg1NzQ1WqCCEzowggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMIIGtDCCBJygAwIBAgIQDcesVwX/IZku
# QEMiDDpJhjANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQD
# ExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjUwNTA3MDAwMDAwWhcNMzgw
# MTE0MjM1OTU5WjBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIElu
# Yy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJT
# QTQwOTYgU0hBMjU2IDIwMjUgQ0ExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAtHgx0wqYQXK+PEbAHKx126NGaHS0URedTa2NDZS1mZaDLFTtQ2oRjzUX
# MmxCqvkbsDpz4aH+qbxeLho8I6jY3xL1IusLopuW2qftJYJaDNs1+JH7Z+QdSKWM
# 06qchUP+AbdJgMQB3h2DZ0Mal5kYp77jYMVQXSZH++0trj6Ao+xh/AS7sQRuQL37
# QXbDhAktVJMQbzIBHYJBYgzWIjk8eDrYhXDEpKk7RdoX0M980EpLtlrNyHw0Xm+n
# t5pnYJU3Gmq6bNMI1I7Gb5IBZK4ivbVCiZv7PNBYqHEpNVWC2ZQ8BbfnFRQVESYO
# szFI2Wv82wnJRfN20VRS3hpLgIR4hjzL0hpoYGk81coWJ+KdPvMvaB0WkE/2qHxJ
# 0ucS638ZxqU14lDnki7CcoKCz6eum5A19WZQHkqUJfdkDjHkccpL6uoG8pbF0LJA
# QQZxst7VvwDDjAmSFTUms+wV/FbWBqi7fTJnjq3hj0XbQcd8hjj/q8d6ylgxCZSK
# i17yVp2NL+cnT6Toy+rN+nM8M7LnLqCrO2JP3oW//1sfuZDKiDEb1AQ8es9Xr/u6
# bDTnYCTKIsDq1BtmXUqEG1NqzJKS4kOmxkYp2WyODi7vQTCBZtVFJfVZ3j7OgWmn
# hFr4yUozZtqgPrHRVHhGNKlYzyjlroPxul+bgIspzOwbtmsgY1MCAwEAAaOCAV0w
# ggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFO9vU0rp5AZ8esrikFb2
# L9RJ7MtOMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RH
# NC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIw
# CwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQAXzvsWgBz+Bz0RdnEwvb4L
# yLU0pn/N0IfFiBowf0/Dm1wGc/Do7oVMY2mhXZXjDNJQa8j00DNqhCT3t+s8G0iP
# 5kvN2n7Jd2E4/iEIUBO41P5F448rSYJ59Ib61eoalhnd6ywFLerycvZTAz40y8S4
# F3/a+Z1jEMK/DMm/axFSgoR8n6c3nuZB9BfBwAQYK9FHaoq2e26MHvVY9gCDA/JY
# sq7pGdogP8HRtrYfctSLANEBfHU16r3J05qX3kId+ZOczgj5kjatVB+NdADVZKON
# /gnZruMvNYY2o1f4MXRJDMdTSlOLh0HCn2cQLwQCqjFbqrXuvTPSegOOzr4EWj7P
# tspIHBldNE2K9i697cvaiIo2p61Ed2p8xMJb82Yosn0z4y25xUbI7GIN/TpVfHIq
# Q6Ku/qjTY6hc3hsXMrS+U0yy+GWqAXam4ToWd2UQ1KYT70kZjE4YtL8Pbzg0c1ug
# MZyZZd/BdHLiRu7hAWE6bTEm4XYRkA6Tl4KSFLFk43esaUeqGkH/wyW4N7Oigizw
# JWeukcyIPbAvjSabnf7+Pu0VrFgoiovRDiyx3zEdmcif/sYQsfch28bZeUz2rtY/
# 9TCA6TD8dC3JE3rYkrhLULy7Dc90G6e8BlqmyIjlgp2+VqsS9/wQD7yFylIz0scm
# bKvFoW2jNrbM1pD2T7m3XDCCBY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFow
# DQYJKoZIhvcNAQEMBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNl
# cnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIz
# NTk1OVowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcG
# A1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3Rl
# ZCBSb290IEc0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2je
# u+RdSjwwIjBpM+zCpyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bG
# l20dq7J58soR0uRf1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBE
# EC7fgvMHhOZ0O21x4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/N
# rDRAX7F6Zu53yEioZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A
# 2raRmECQecN4x7axxLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8
# IUzUvK4bA3VdeGbZOjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfB
# aYh2mHY9WV1CdoeJl2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaa
# RBkrfsCUtNJhbesz2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZi
# fvaAsPvoZKYz0YkH4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXe
# eqxfjT/JvNNBERJb5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g
# /KEexcCPorF+CiaZ9eRpL5gdLfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB
# /wQFMAMBAf8wHQYDVR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1Ud
# IAQKMAgwBgYEVR0gADANBgkqhkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22
# Ftf3v1cHvZqsoYcs7IVeqRq7IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih
# 9/Jy3iS8UgPITtAq3votVs/59PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYD
# E3cnRNTnf+hZqPC/Lwum6fI0POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c
# 2PR3WlxUjG/voVA9/HYJaISfb8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88n
# q2x2zm8jLfR+cWojayL/ErhULSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5
# lDGCA3wwggN4AgEBMH0waTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMQIQCoDvGEuN8QWC0cR2p5V0aDANBglg
# hkgBZQMEAgEFAKCB0TAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZI
# hvcNAQkFMQ8XDTI1MDYxODE4NTc0NVowKwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU
# 3WIwrIYKLTBr2jixaHlSMAf7QX4wLwYJKoZIhvcNAQkEMSIEIPqVlfanSp3FmK5x
# oR9rKhicQV8ZWm/8G7Q4AjwqXHqkMDcGCyqGSIb3DQEJEAIvMSgwJjAkMCIEIEqg
# P6Is11yExVyTj4KOZ2ucrsqzP+NtJpqjNPFGEQozMA0GCSqGSIb3DQEBAQUABIIC
# AAHJo0S6XghNg0Txws7FfwiUhgY7xJeKnTqcON32IFL1Lumty3LzMTE5gedG+puD
# Yf4Qemec9ikmIpga+f8b2Z1dFcEvu7aA4NPm/zzINSZ/7/JszBYzaK50mummbKoK
# wFMtaOy2vfbLIaIqmxlu3vyTvHdHdU0JsRz1LpFr6Y9YiuAawAh70I9atO3wDMJ4
# U31n9pnGGBIWyZL0exZF+BnSgxse23ldOiOS3dtT4b+nJuwBxizUmo45RFz0QW2n
# PvoA+IkM9dHXpsZVVaGquMWWMFhiPUl+U+2jadRR0pXdxIGPf/k1NZjH6myXdgSB
# azpKtyAsIVIeYdx6JtWdw45HJnCtIutf7/GOom5K/6ZVjr+QxgguPssDOynItglj
# ozP5NE02AjIsRLn9tKJE0WKtgSJHsrQZqhzdz2KI13vWitHQk1XJmSU5MkiZhUjC
# aVtS68HfCVArR/L5ggdcfHjaLfzQiRBSW4nnqSN7hr9+YIPT+OoYk3jlarLc19+C
# w+DVScspgjKLdxGX3TUgEdLSKYtKapsK0Q9tnlrG73F4KEDLaFyVFODj62xa2BGD
# 2tW6qp8iPHn++R930JNVA/5Fs3o1FyTf9hY13aa7366NRReeIoUb5NjnE42ZoX1G
# MQ01MBu7yfY8y9pRpyxwJIKuct2l3vgU3f53NFQPzxBt
# SIG # End signature block
