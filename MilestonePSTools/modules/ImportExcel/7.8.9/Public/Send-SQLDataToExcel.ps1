function Send-SQLDataToExcel {
    [CmdletBinding(DefaultParameterSetName="none")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification="Allowed to use DBSessions Global variable from GETSQL Module")]

    param (
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",  Mandatory=$true)]
        $Connection,
        [Parameter(ParameterSetName="ExistingSession", Mandatory=$true)]
        $Session,
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [switch]$MsSqlServer,
        [Parameter(ParameterSetName="SQLConnection")]
        [String]$DataBase,
        [Parameter(ParameterSetName="SQLConnection",   Mandatory=$true)]
        [Parameter(ParameterSetName="ODBCConnection",  Mandatory=$true)]
        [Parameter(ParameterSetName="ExistingSession", Mandatory=$true)]
        [string]$SQL,
        [int]$QueryTimeout,
        [Parameter(ParameterSetName="Pre-FetchedData", Mandatory=$true)]
        [System.Data.DataTable]$DataTable,
        [switch]$Force
    )
#Import the parameters from Export-Excel, we will pass InputObject, and we have the common parameters so exclude those,
#and re-write the [Parmameter] attribute on each one to avoid parameterSetName here competing with the settings in Export excel.
#The down side of this that impossible parameter combinations won't be filtered out and need to be caught later.
    DynamicParam {
        $ParameterAttribute  =                        "System.Management.Automation.ParameterAttribute"
        $RuntimeDefinedParam =                        "System.Management.Automation.RuntimeDefinedParameter"
        $paramDictionary     =    New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
        $attributeCollection =    New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
        $attributeCollection.Add((New-Object -TypeName $ParameterAttribute -Property @{ ParameterSetName = "__AllParameterSets" ;Mandatory = $false}))
        foreach ($P in (Get-Command -Name Export-Excel).Parameters.values.where({$_.name -notmatch 'Verbose|Debug|Action$|Variable$|Buffer$|TargetData$|InputObject$'}))  {
            $paramDictionary.Add($p.Name, (New-Object -TypeName $RuntimeDefinedParam -ArgumentList $p.name, $p.ParameterType, $attributeCollection ) )
        }
        return $paramDictionary
    }
    process {
      #region Dynamic params mean we can get passed parameter combination Export-Excel will reject, so throw here, rather than get data and then have Export-Excel error.
        if ($PSBoundParameters.Path -and $PSBoundParameters.ExcelPackage) {
            throw 'Parameter error: you cannot specify both a path and an Excel Package.'
            return
        }
        if ($PSBoundParameters.AutoFilter -and ($PSBoundParameters.TableName -or $PSBoundParameters.TableStyle)) {
            Write-Warning "Tables are automatically auto-filtered, -AutoFilter will be ignored"
            $null = $PSBoundParameters.Remove('AutoFilter')
        }
      #endregion
      #region if we were either given a session object or a connection string (& optionally -MsSqlServer) make sure we can connect
        try {
            #If we got -MsSqlServer, create a SQL connection, if we didn't but we got -Connection create an ODBC connection
            if     ($MsSqlServer -and $Connection) {
                if ($Connection -notmatch '=') {$Connection = "server=$Connection;trusted_connection=true;timeout=60"}
                $Session     = New-Object -TypeName System.Data.SqlClient.SqlConnection  -ArgumentList $Connection
                if ($Session.State -ne 'Open') {$Session.Open()}
                if ($DataBase) {$Session.ChangeDatabase($DataBase) }
            }
            elseif ($Connection)  {
                $Session     = New-Object -TypeName System.Data.Odbc.OdbcConnection      -ArgumentList $Connection ; $Session.ConnectionTimeout = 30
            }
        }
        catch {
            Write-Warning "An Error occured trying to connect to $Connection, the error was $([Environment]::NewLine + $_.Exception.InnerException))"
        }
        if ($Session -is [String] -and $Global:DbSessions[$Session]) {$Session = $Global:DbSessions[$Session]}
      #endregion
      #region we may have been given a table, but if there is a db session to connect to, send the query
        if     ($Session) {
            try {
                #If the session a SQL one make a SQL DataAdapter, otherwise make an ODBC one
                if ($Session.GetType().name -match "SqlConnection") {
                    $dataAdapter = New-Object -TypeName System.Data.SqlClient.SqlDataAdapter -ArgumentList (
                                New-Object -TypeName System.Data.SqlClient.SqlCommand     -ArgumentList $SQL, $Session)
                }
                else {
                    $dataAdapter = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter     -ArgumentList (
                                New-Object -TypeName System.Data.Odbc.OdbcCommand         -ArgumentList $SQL, $Session )
                }
                if ($QueryTimeout) {$dataAdapter.SelectCommand.CommandTimeout = $QueryTimeout}

                #Both adapter types output the same kind of table, create one and fill it from the adapter
                $DataTable       = New-Object -TypeName System.Data.DataTable
                $rowCount        = $dataAdapter.fill($dataTable)
                Write-Verbose -Message "Query returned $rowCount row(s)"
            }
            catch {
                Write-Warning "An Error occured trying to run the query, the error was $([Environment]::NewLine + $_.Exception.InnerException))"
            }
        }
      #endregion
      #region send the table to Excel
        #remove parameters which relate to querying SQL, leaving the ones used by Export-Excel
        'Connection' , 'Database'  , 'Session' , 'MsSqlServer' , 'SQL'  , 'DataTable'  , 'QueryTimeout' , 'Force' |
                ForEach-Object {$null = $PSBoundParameters.Remove($_) }
        #if force was specified export even if there are no rows. If there are no columns, the query failed and export "null" if forced
        if     ($DataTable.Rows.Count) {
             Export-Excel  @PSBoundParameters -InputObject $DataTable
        }
        elseif ($Force -and $DataTable.Columns.Count) {
            Write-Warning -Message "Zero rows returned, and -Force was specified, sending empty table to Excel."
            Export-Excel  @PSBoundParameters -InputObject $DataTable
        }
        elseif ($Force) {
            Write-Warning -Message "-Force was specified but there is no data to send."
            Export-Excel  @PSBoundParameters -InputObject $null
        }
        else   {Write-Warning -Message 'There is no Data to insert, and -Force was not specified.' }
      #endregion
      #If we were passed a connection and opened a session,  close that session.
        if     ($Connection)  {$Session.close() }
    }
}
# SIG # Begin signature block
# MIIuzgYJKoZIhvcNAQcCoIIuvzCCLrsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA6Nt0aITz86Mgn
# n17dTcHkYjwTBI2A7vSl8y4o4a22S6CCE24wggVyMIIDWqADAgECAhB2U/6sdUZI
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
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDVicOkH3aFXKgnC4Tf1nYEzBaWrnsi
# /L7GWQdm4XfDAjA4BgorBgEEAYI3AgEMMSowKKACgAChIoAgaHR0cHM6Ly93d3cu
# bWlsZXN0b25lcHN0b29scy5jb20wDQYJKoZIhvcNAQEBBQAEggIAYpLxdLjiCIsp
# hIK/OYPVfMqQdv6Zi+HGg9dqujDbCPxJQFNI8Zrt7BiTgSwgaNNdHVxxJXxpkg4f
# pPIMZYPAM1Q1nrUqi7bp6JQHc/uuB0MXsSmxzc9R8v6edYCfh+d/gbKf9wMkndkp
# DdLGJbOUTDiySFvo+w7xN6t48WrTJ+ftOWe5w3ApaWjh4VS1oHBaGiAcrctG4sjR
# CZBcTE8IclleF4ohXdqUuLmf5bgQ4ISYc+HvVJ619p5H1AyvY3RSq1WJ1VknmGrd
# x0ZUiuLTTZesQahaWDRaKPgaaphuQA55LHfrREevkU2ZXZQr5EXYY3JepGN7TGMw
# f4UP2OPP9liTbTqxlfEmrLarltH2QpL73p0VEHba0qZHtMYyYBLvRDp39wi0m194
# v4K9ZGj8me46OGFkteeiplTJr7V+tBuIIx0sFd9B8pslGPkMSgUx3EKWJrI9esfL
# wNoF8+p6MF6IomyMMTZ/eBGrLdBg6AMRtKfiJk+vkbg+P57cm9qkQRoWAXAi7mST
# 5+ypsFMgXqZ9r3g9NJ5tKlRcIpEGrAYFWzmEgc63U4GB0eKuURqLVNeGJadxJXOW
# nxlz0Z6HO6RwEG5gTMEUxtEcWzRO34eggHpxLi72qIspMlBsobaR0Up5Gs2kFd8+
# pSxw1Oic2QckZa345by/64TTcKGlEwehghd3MIIXcwYKKwYBBAGCNwMDATGCF2Mw
# ghdfBgkqhkiG9w0BBwKgghdQMIIXTAIBAzEPMA0GCWCGSAFlAwQCAQUAMHgGCyqG
# SIb3DQEJEAEEoGkEZzBlAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# g2lfz05IlKD4CVRkp0oZigd3JKiM2uYGdT2DdTOQ1iwCEQCH4iJ1jalXO1ikcoH2
# 6B6AGA8yMDI1MDYxODE4NTc0OFqgghM6MIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC
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
# SIb3DQEJBTEPFw0yNTA2MTgxODU3NDhaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYE
# FN1iMKyGCi0wa9o4sWh5UjAH+0F+MC8GCSqGSIb3DQEJBDEiBCAZ7fimVhIGOmSd
# OSkl7x84SejzNxqDGAfoioSJHqP4fzA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCBK
# oD+iLNdchMVck4+CjmdrnK7Ksz/jbSaaozTxRhEKMzANBgkqhkiG9w0BAQEFAASC
# AgBs9c+kg67m/CAgJtIGTo+fg6SDkr3YFNRdhXzzkyfK3Ivs1enY/bjsFPVDgFMY
# /THlQrK/TIO/xmYHWzpRpjGxdb3OSZ8ivKX4tolT68SuR09GIkOztymT6lfp+T/C
# sDf+kDnHXrZAHhywjctSyCo57teLO2Ld1wUjzikUVRV7SybRHyVDAXF28Jqi7gYH
# KkuDrrhdyMX93q2pafhIEOggWEgl4d2FSVdI47MMWWTX4axQcTlnHUJOnF+I8bA+
# rw1opn+Od5kKy38fArq7iqlskwQryQvF2MZ5yb+acaSCBHanP+Ov8DyBX/AGVAYY
# FKBpsyGDAQKATRqmdZSJBrYxosS5DESjq/EUUg6YZUm3/sxNkqv4en3wMzV/wOgH
# 4n+xgAsZKy0wBjyH+EzQ+5O5mIw8HzqK130Gye733es5MKGgUglnYVKh5yuUZ0r8
# w7uxrx/L5+ncEjh/J28tYiYmZuO88A7iu9Jv58fhE6875TVePPEhvDimyCsayYOQ
# fKz9XmA20LTSkszzRspForYprfj7KnrrNpttNh5+8rHxuwYrSZgLYJEFKBEJTt1o
# hKWymPoQJz9ZVIyO5ZZMcXGMQNSM95eS5WRa7auj3CVDtBHvXEFQR6Q4CCrA+ziM
# a3eotUPUgQooVrIZjV943fGAwn2uIVNVjxLZ/t7SGvJg8w==
# SIG # End signature block
