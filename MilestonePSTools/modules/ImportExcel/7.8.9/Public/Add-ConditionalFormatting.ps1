function Add-ConditionalFormatting {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias("Range")]
        $Address ,
        [OfficeOpenXml.ExcelWorksheet]$Worksheet ,
        [Parameter(Mandatory = $true, ParameterSetName = "NamedRule", Position = 1)]
        [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType]$RuleType ,
        [Parameter(ParameterSetName = "NamedRule")]
        [Alias("ForegroundColour","FontColor")]
        $ForegroundColor,
        [Parameter(Mandatory = $true, ParameterSetName = "DataBar")]
        [Alias("DataBarColour")]
        $DataBarColor,
        [Parameter(Mandatory = $true, ParameterSetName = "ThreeIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]$ThreeIconsSet,
        [Parameter(Mandatory = $true, ParameterSetName = "FourIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting4IconsSetType]$FourIconsSet,
        [Parameter(Mandatory = $true, ParameterSetName = "FiveIconSet")]
        [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting5IconsSetType]$FiveIconsSet,
        [Parameter(ParameterSetName = "NamedRule")]
        [Parameter(ParameterSetName = "ThreeIconSet")]
        [Parameter(ParameterSetName = "FourIconSet")]
        [Parameter(ParameterSetName = "FiveIconSet")]
        [switch]$Reverse,
        [switch]$ShowIconOnly,
        [Parameter(ParameterSetName = "NamedRule",Position = 2)]
        $ConditionValue,
        [Parameter(ParameterSetName = "NamedRule",Position = 3)]
        $ConditionValue2,
        [Parameter(ParameterSetName = "NamedRule")]
        $BackgroundColor,
        [Parameter(ParameterSetName = "NamedRule")]
        [OfficeOpenXml.Style.ExcelFillStyle]$BackgroundPattern = [OfficeOpenXml.Style.ExcelFillStyle]::None ,
        [Parameter(ParameterSetName = "NamedRule")]
        $PatternColor,
        [Parameter(ParameterSetName = "NamedRule")]
        $NumberFormat,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Bold,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Italic,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$Underline,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$StrikeThru,
        [Parameter(ParameterSetName = "NamedRule")]
        [switch]$StopIfTrue,
        [int]$Priority,
        [switch]$PassThru
    )

    #Allow conditional formatting to work like Set-ExcelRange (with single ADDRESS parameter), split it to get worksheet and range of cells.
    if ($Address -is [OfficeOpenXml.Table.ExcelTable]) {
            $Worksheet = $Address.Address.Worksheet
            $Address   = $Address.Address.Address
    }
    elseif  ($Address.Address -and $Address.Worksheet -and -not $Worksheet) { #Address is a rangebase or similar
        $Worksheet = $Address.Worksheet[0]
        $Address   = $Address.Address
    }
    elseif ($Address -is [String] -and $Worksheet -and $Worksheet.Names[$Address] ) { #Address is the name of a named range.
        $Address = $Worksheet.Names[$Address].Address
    }
    if (($Address -is [OfficeOpenXml.ExcelRow]    -and -not $Worksheet) -or
        ($Address -is [OfficeOpenXml.ExcelColumn] -and -not $Worksheet) ){  #EPPLUs Can't get the worksheet object from a row or column object, so bail if that was tried
        Write-Warning -Message "Add-ConditionalFormatting does not support Row or Column objects as an address; use a worksheet and/or specify 'R:R' or 'C:C' instead. "; return
    }
    elseif ($Address -is [OfficeOpenXml.ExcelRow]) {  #But if we have a column or row object and a worksheet (I don't know *why*) turn them into a string for the range
            $Address = "$($Address.Row):$($Address.Row)"
    }
    elseif ($Address -is [OfficeOpenXml.ExcelColumn]) {
        $Address = (New-Object 'OfficeOpenXml.ExcelAddress' @(1, $address.ColumnMin, 1, $address.ColumnMax)).Address -replace '1',''
        if ($Address -notmatch ':') {$Address = "$Address`:$Address"}
    }
    if ( $Address -is [string] -and $Address -match "!") {$Address = $Address -replace '^.*!',''}
    #By this point we should have a worksheet object whose ConditionalFormatting collection we will add to. If not, bail.
    if (-not $worksheet -or $Worksheet -isnot [OfficeOpenXml.ExcelWorksheet]) {write-warning "You need to provide a worksheet object." ; return}
    #region create a rule of the right type
    if     ($RuleType -match 'IconSet$') {Write-warning -Message "You cannot configure a Icon-Set rule in this way; please use -$RuleType <SetName>." ; return}
    if ($PSBoundParameters.ContainsKey("DataBarColor"  )      ) {if ($DataBarColor -is [string]) {$DataBarColor = [System.Drawing.Color]::$DataBarColor }
                                                                     $rule =  $Worksheet.ConditionalFormatting.AddDatabar(     $Address , $DataBarColor )
    }
    elseif ($PSBoundParameters.ContainsKey("ThreeIconsSet" )      ) {$rule =  $Worksheet.ConditionalFormatting.AddThreeIconSet($Address , $ThreeIconsSet)}
    elseif ($PSBoundParameters.ContainsKey("FourIconsSet"  )      ) {$rule =  $Worksheet.ConditionalFormatting.AddFourIconSet( $Address , $FourIconsSet )}
    elseif ($PSBoundParameters.ContainsKey("FiveIconsSet"  )      ) {$rule =  $Worksheet.ConditionalFormatting.AddFiveIconSet( $Address , $FiveIconsSet )}
    else                                                            {$rule = ($Worksheet.ConditionalFormatting)."Add$RuleType"($Address )                }
    If($ShowIconOnly) {
        $rule.ShowValue = $false
    }
    if     ($Reverse)  {
            if     ($rule.type -match 'IconSet$'   )                {$rule.reverse = $true}
            elseif ($rule.type -match 'ColorScale$')                {$temp =$rule.LowValue.Color ; $rule.LowValue.Color = $rule.HighValue.Color; $rule.HighValue.Color = $temp}
            else   {Write-Warning -Message "-Reverse was ignored because $RuleType does not support it."}
    }
    #endregion
    #region set the rule conditions
    #for lessThan/GreaterThan/Equal/Between conditions make sure that strings are wrapped in quotes. Formulas should be passed with = which will be stripped.
    if     ($RuleType -match "Than|Equal|Between" ) {
        if  ($PSBoundParameters.ContainsKey("ConditionValue" )) {
                $number = $Null
                #if the condition type is not a value type, but parses as a number, make it the number
                if ($ConditionValue -isnot [System.ValueType] -and [Double]::TryParse($ConditionValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number) ) {
                         $ConditionValue  = $number
                } #else if it is not a value type, or a formula, or wrapped in quotes, wrap it in quotes.
                elseif (($ConditionValue -isnot [System.ValueType])-and ($ConditionValue  -notmatch '^=') -and ($ConditionValue  -notmatch '^".*"$') ) {
                         $ConditionValue  = '"' + $ConditionValue +'"'
                }
        }
        if  ($PSBoundParameters.ContainsKey("ConditionValue2")) {
                $number = $Null
                if ($ConditionValue -isnot [System.ValueType] -and [Double]::TryParse($ConditionValue2, [System.Globalization.NumberStyles]::Any, [System.Globalization.NumberFormatInfo]::CurrentInfo, [Ref]$number) ) {
                         $ConditionValue2 = $number
                }
                elseif (($ConditionValue -isnot [System.ValueType]) -and ($ConditionValue2 -notmatch '^=') -and ($ConditionValue2 -notmatch '^".*"$') ) {
                         $ConditionValue2  = '"' + $ConditionValue2 + '"'
                }
        }
    }
    #But we don't usually want quotes round containstext | beginswith type rules. Can't be Certain they need to be removed, so warn the user their condition might be wrong
    if     ($RuleType -match "Text|With" -and $ConditionValue -match '^".*"$'  ) {
            Write-Warning -Message "The condition will look for the quotes at the start and end."
    }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "Top|Bottom"                          ) {$rule.Rank      = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "StdDev"                             ) {$rule.StdDev    = $ConditionValue }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "Than|Equal|Expression"              ) {$rule.Formula   = ($ConditionValue  -replace '^=','') }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $RuleType -match "Text|With"                          ) {$rule.Text      = ($ConditionValue  -replace '^=','') }
    if     ($PSBoundParameters.ContainsKey("ConditionValue" ) -and
            $PSBoundParameters.ContainsKey("ConditionValue2") -and
            $RuleType -match "Between"                            ) {
                                                                     $rule.Formula   = ($ConditionValue  -replace '^=','');
                                                                     $rule.Formula2  = ($ConditionValue2 -replace '^=','')
    }
    if     ($PSBoundParameters.ContainsKey("StopIfTrue")          ) {$rule.StopIfTrue = $StopIfTrue }
    if     ($PSBoundParameters.ContainsKey("Priority")            ) {$rule.Priority   = $Priority }
    #endregion
    #region set the rule format
    if     ($PSBoundParameters.ContainsKey("NumberFormat"     )   ) {$rule.Style.NumberFormat.Format        = (Expand-NumberFormat  $NumberFormat)             }
    if     ($Underline                                            ) {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::Single }
    elseif ($PSBoundParameters.ContainsKey("Underline"        )   ) {$rule.Style.Font.Underline             = [OfficeOpenXml.Style.ExcelUnderLineType]::None   }
    if     ($PSBoundParameters.ContainsKey("Bold"             )   ) {$rule.Style.Font.Bold                  = [boolean]$Bold       }
    if     ($PSBoundParameters.ContainsKey("Italic"           )   ) {$rule.Style.Font.Italic                = [boolean]$Italic     }
    if     ($PSBoundParameters.ContainsKey("StrikeThru"       )   ) {$rule.Style.Font.Strike                = [boolean]$StrikeThru }
    if     ($PSBoundParameters.ContainsKey("ForeGroundColor"  )   ) {if ($ForeGroundColor -is [string])      {$ForeGroundColor = [System.Drawing.Color]::$ForeGroundColor }
                                                                     $rule.Style.Font.Color.color           = $ForeGroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundColor"  )   ) {if ($BackgroundColor -is [string])      {$BackgroundColor = [System.Drawing.Color]::$BackgroundColor }
                                                                     $rule.Style.Fill.BackgroundColor.color = $BackgroundColor     }
    if     ($PSBoundParameters.ContainsKey("BackgroundPattern")   ) {$rule.Style.Fill.PatternType           = $BackgroundPattern   }
    if     ($PSBoundParameters.ContainsKey("PatternColor"     )   ) {if ($PatternColor -is [string])         {$PatternColor = [System.Drawing.Color]::$PatternColor }
                                                                     $rule.Style.Fill.PatternColor.color    = $PatternColor        }
    #endregion
    #Allow further tweaking by returning the rule, if passthru specified
    if     ($Passthru)  {$rule}
}
# SIG # Begin signature block
# MIIuzgYJKoZIhvcNAQcCoIIuvzCCLrsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOivkEUnWI0z0u
# 9koDaW8BuDz28IL8dL1EpWRD/u1L2KCCE24wggVyMIIDWqADAgECAhB2U/6sdUZI
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
# AkA8F4fhc3Yxghq2MIIasgIBATBpMFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBH
# bG9iYWxTaWduIG52LXNhMS8wLQYDVQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29k
# ZVNpZ25pbmcgQ0EgMjAyMAIMcGmCcIM07v5TgMt1MA0GCWCGSAFlAwQCAQUAoIGk
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC0eL3AvdjndhuJLRZm7s+b02nNCBDA
# 8clGHR8bToTaJDA4BgorBgEEAYI3AgEMMSowKKACgAChIoAgaHR0cHM6Ly93d3cu
# bWlsZXN0b25lcHN0b29scy5jb20wDQYJKoZIhvcNAQEBBQAEggIAW56HD5iYCMbw
# b0NKgL7Bl7lqHhiTfxLANngtaQ/PCmVzq9OTbOqjoZbM3hr1weVxlPEp/XX5j30p
# 9gI7h7hCq2vLxOzqsaqVTl8K8hgUV2Q208uHUaNwrl19fGUEgZuWTXgkdBB66fKW
# HOMZT1BgdFIMHe0HjseleuWAcPtYinazmhiR2tnc/xIpk/bRX7mNz32PYlVY9PNa
# 63RWmMh4casDbkJ/OsMsoCOMRq2Rsx75gFAmE5LfkVAF1xt0zgni2ZmydXHo6fuI
# SgTRnAAwzx9uiV9WZ3zabxn4tAsmIB/+3AoGzLtsivALu8S6JdIU51RF5MRiidRh
# u1Wz99qsvA0Pgwlaf0frUb0/+Sjutw9aDJgULy6mxDmAZGMIc0Z15jzL0dbpZanT
# 2ZvH4LFNyp96TcwuikRv7TYhrTyXceweQLRTYrK9wE2nXtH6qBO3kp3nvUf8IvEP
# FkaRxGH91HdGCEuwxrEm97DkWymsZfnWZCXbbhEt7vy+YBIzWR/I4IpUha4ANGnQ
# Y5Qhdb1fnHpT0VxLTdkTkz2cwa2NlcWmhWgirGxMn6c1uZ8FLs5CjSmL+wd/FWJI
# tDbOnMz/00Mo3qD+hBFLzZlkD71LDeIkdTqAwMMn+FsFNEWsTlWd6YmZMHRqwM7N
# iP/t1reIs8bSexEmUjoB7+/G8yCJX7qhghd3MIIXcwYKKwYBBAGCNwMDATGCF2Mw
# ghdfBgkqhkiG9w0BBwKgghdQMIIXTAIBAzEPMA0GCWCGSAFlAwQCAQUAMHgGCyqG
# SIb3DQEJEAEEoGkEZzBlAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# q1TcVXpnUm/FAxMOOgbwzpzj1XYlym/kgnM8CCATdtgCEQConMgt2XiYsNJqMBCB
# eRWQGA8yMDI1MDYxODE4NTc0NVqgghM6MIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC
# 0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGlt
# ZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAw
# MFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1l
# c3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA
# 69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6w
# W2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00
# Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOM
# A3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmot
# uQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1O
# pbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeH
# VZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1r
# oSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSURO
# wnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW0
# 0aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGV
# MIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM
# 6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMC
# B4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKG
# UWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSg
# UqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRU
# aW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAI
# BgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcE
# ua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/Ym
# RDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8
# AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/E
# ABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQ
# VTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gV
# utDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85
# EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hg
# gt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJ
# gKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLv
# UxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7P
# OGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCBrQwggScoAMCAQICEA3HrFcF/yGZ
# LkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UE
# AxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoXDTM4
# MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBS
# U0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNqEY81
# FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fkHUil
# jNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EEbkC9
# +0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8NF5v
# p7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUUFREm
# DrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP9qh8
# SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKWxdCy
# QEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespYMQmU
# iote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrPV6/7
# umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+zoFp
# p4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK4pBW
# 9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZxML2+
# C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97frPBtI
# j+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+NMvE
# uBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYAgwPy
# WLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA1WSj
# jf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+BFo+
# z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06VXxy
# KkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284NHNb
# oDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDezooIs
# 8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM9q7W
# P/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpSM9LH
# JmyrxaFtoza2zNaQ9k+5t1wwggWNMIIEdaADAgECAhAOmxiO+dAt5+/bUOIIQBha
# MA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0zMTExMDky
# MzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/mkHNo
# 3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4FutW
# xpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMylNEQ
# RBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq868nP
# zaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe3I6P
# gNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMqbpEB
# fCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxGj2T3
# wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORFJYm2
# mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhElRp2
# Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0viastkF1
# 3nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LWRV+d
# IPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAPBgNVHRMB
# Af8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYweQYIKwYB
# BQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDARBgNV
# HSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4X6CjdBs9
# thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc47/T/gLn4
# offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0mcudT6cG
# AxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk1T3osdz9
# HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLBsln1VWvP
# J6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGpn1eqXiji
# uZQxggN8MIIDeAIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBp
# bmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJ
# YIZIAWUDBAIBBQCggdEwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwGCSqG
# SIb3DQEJBTEPFw0yNTA2MTgxODU3NDVaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYE
# FN1iMKyGCi0wa9o4sWh5UjAH+0F+MC8GCSqGSIb3DQEJBDEiBCBfBaHXUysi/Qss
# cus/+gDbzqQaJI0g/OmJaujoNz5zQjA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCBK
# oD+iLNdchMVck4+CjmdrnK7Ksz/jbSaaozTxRhEKMzANBgkqhkiG9w0BAQEFAASC
# AgA4I24lpNHTtXrle9g7LCAU5TmsRmBr1CVFdxDH+lwjewN7gxlOnmaV7E+HNElR
# GwZVUgu/m/mWQK/qbkkqTgsNhdPLKeQ58qzhJI5tfP+WsXhsmuICKRUnXd5z02rl
# V8LyLV9ohOCtdXX2ityC8mwWO+rWS2K72K7HYTQSV2MJoYjS4n90eMZ2Pu7FfJj3
# StiUbESG7aaLCUeftFqmKaFq2CPLGHMEVOaS1VqfcSc2gyE66uYXsJwI7+/G69C9
# J4veU94omhnntXt8Vb0HblZIi1QoJSbKzxjRV0SKK68ltd0OrCVdTB11J4TUWeYq
# g7bEMsxZp555tBNM78u7+/lSeEFXFfrfMS2i6/DwFxI9QQarP+A9lrMxQgRgim5u
# ZUIdyTMPRN+cg4YKqz1UO+cfGrGUu9HnoLzvWVuqYQc3olkGKEnhAC2V2rViumsO
# HBqoQHvqt38+gEQIvNgqfvpGSuTT0xsT4BZIksVoInAeW7itmDywqUCo86Ud+JPw
# b51YHBSsadNICXTvPmMtQh8yTbTzSGM2t8DMvMIieoM9u23MgxI6Nl+06fl3tPyI
# mw/AOS6kVIvy/s9uXRv3zNDi+wjtoq2FwRSJGjTmiBFwaxrEj7NAdeA70EP+vFVN
# kN0/Pi0hhJr1oDRqMJmP1nHvVOlXxCqdoim7F5LErvlr+A==
# SIG # End signature block
