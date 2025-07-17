## Start-Demo.ps1
##################################################################################################
## This is an overhaul of Jeffrey Snover's original Start-Demo script by Joel "Jaykul" Bennett
##
## I've switched it to using ReadKey instead of ReadLine (you don't have to hit Enter each time)
## As a result, I've changed the names and keys for a lot of the operations, so that they make
## sense with only a single letter to tell them apart (sorry if you had them memorized).
##
## I've also been adding features as I come across needs for them, and you'll contribute your
## improvements back to the PowerShell Script repository as well.
##################################################################################################
## Revision History (version 3.3)
## 3.3.3 Fixed:    Script no longer says "unrecognized key" when you hit shift or ctrl, etc.
##       Fixed:    Blank lines in script were showing as errors (now printed like comments)
## 3.3.2 Fixed:    Changed the "x" to match the "a" in the help text
## 3.3.1 Fixed:    Added a missing bracket in the script
## 3.3 - Added:    Added a "Clear Screen" option
##     - Added:    Added a "Rewind" function (which I'm not using much)
## 3.2 - Fixed:    Put back the trap { continue; }
## 3.1 - Fixed:    No Output when invoking Get-Member (and other cmdlets like it???)
## 3.0 - Fixed:    Commands which set a variable, like: $files = ls
##     - Fixed:    Default action doesn't continue
##     - Changed:  Use ReadKey instead of ReadLine
##     - Changed:  Modified the option prompts (sorry if you had them memorized)
##     - Changed:  Various time and duration strings have better formatting
##     - Enhance:  Colors are settable: prompt, command, comment
##     - Added:    NoPauseAfterExecute switch removes the extra pause
##                 If you set this, the next command will be displayed immediately
##     - Added:    Auto Execute mode (FullAuto switch) runs the rest of the script
##                 at an automatic speed set by the AutoSpeed parameter (or manually)
##     - Added:    Automatically append an empty line to the end of the demo script
##                 so you have a chance to "go back" after the last line of you demo
##################################################################################################
##
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Correct and desirable usage')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingInvokeExpression', '', Justification='Correct and desirable usage')]
param(
  $file=".\demo.txt",
  [int]$command=0,
  [System.ConsoleColor]$promptColor="Yellow",
  [System.ConsoleColor]$commandColor="White",
  [System.ConsoleColor]$commentColor="Green",
  [switch]$FullAuto,
  [int]$AutoSpeed = 3,
  [switch]$NoPauseAfterExecute
)

$RawUI = $Host.UI.RawUI
$hostWidth = $RawUI.BufferSize.Width

# A function for reading in a character
function Read-Char() {
  $_OldColor = $RawUI.ForeGroundColor
  $RawUI.ForeGroundColor = "Red"
  $inChar=$RawUI.ReadKey("IncludeKeyUp")
  # loop until they press a character, so Shift or Ctrl, etc don't terminate us
  while($inChar.Character -eq 0){
    $inChar=$RawUI.ReadKey("IncludeKeyUp")
  }
  $RawUI.ForeGroundColor = $_OldColor
  return $inChar.Character
}

function Rewind($lines, $index, $steps = 1) {
   $started = $index;
   $index -= $steps;
   while(($index -ge 0) -and ($lines[$index].Trim(" `t").StartsWith("#"))){
      $index--
   }
   if( $index -lt 0 ) { $index = $started }
   return $index
}

$file = Resolve-Path $file
while(-not(Test-Path $file)) {
  $file = Read-Host "Please enter the path of your demo script (Crtl+C to cancel)"
  $file = Resolve-Path $file
}

Clear-Host

$_lines = Get-Content $file
# Append an extra (do nothing) line on the end so we can still go back after the last line.
$_lines += "Write-Host 'The End'"
$_starttime = [DateTime]::now
$FullAuto = $false

Write-Host -nonew -back black -fore $promptColor $(" " * $hostWidth)
Write-Host -nonew -back black -fore $promptColor @"
<Demo Started :: $(split-path $file -leaf)>$(' ' * ($hostWidth -(18 + $(split-path $file -leaf).Length)))
"@
Write-Host -nonew -back black -fore $promptColor "Press"
Write-Host -nonew -back black -fore Red " ? "
Write-Host -nonew -back black -fore $promptColor "for help.$(' ' * ($hostWidth -17))"
Write-Host -nonew -back black -fore $promptColor $(" " * $hostWidth)

# We use a FOR and an INDEX ($_i) instead of a FOREACH because
# it is possible to start at a different location and/or jump
# around in the order.
for ($_i = $Command; $_i -lt $_lines.count; $_i++)
{
	# Put the current command in the Window Title along with the demo duration
	$Dur = [DateTime]::Now - $_StartTime
   $RawUI.WindowTitle = "$(if($dur.Hours -gt 0){'{0}h '})$(if($dur.Minutes -gt 0){'{1}m '}){2}s   {3}" -f
                        $dur.Hours, $dur.Minutes, $dur.Seconds, $($_Lines[$_i])

	# Echo out the commmand to the console with a prompt as though it were real
	Write-Host -nonew -fore $promptColor "[$_i]$([char]0x2265) "
	if ($_lines[$_i].Trim(" ").StartsWith("#") -or $_lines[$_i].Trim(" ").Length -le 0) {
		Write-Host -fore $commentColor "$($_Lines[$_i])  "
		continue
	} else {
		Write-Host -nonew -fore $commandColor "$($_Lines[$_i])  "
	}

	if( $FullAuto ) { Start-Sleep $autoSpeed; $ch = [char]13 } else { $ch = Read-Char }
	switch($ch)
	{
		"?" {
			Write-Host -Fore $promptColor @"

Running demo: $file
(n) Next       (p) Previous
(q) Quit       (s) Suspend
(t) Timecheck  (v) View $(split-path $file -leaf)
(g) Go to line by number
(f) Find lines by string
(a) Auto Execute mode
(c) Clear Screen
"@
			$_i-- # back a line, we're gonna step forward when we loop
		}
		"n" { # Next (do nothing)
			Write-Host -Fore $promptColor "<Skipping Line>"
		}
		"p" { # Previous
			Write-Host -Fore $promptColor "<Back one Line>"
			while ($_lines[--$_i].Trim(" ").StartsWith("#")){}
			$_i-- # back a line, we're gonna step forward when we loop
		}
		"a" { # EXECUTE (Go Faster)
			$AutoSpeed = [int](Read-Host "Pause (seconds)")
			$FullAuto = $true;
			Write-Host -Fore $promptColor "<eXecute Remaining Lines>"
			$_i-- # Repeat this line, and then just blow through the rest
		}
		"q" { # Quit
			Write-Host -Fore $promptColor "<Quiting demo>"
			$_i = $_lines.count;
			break;
		}
		"v" { # View Source
			$lines[0..($_i-1)] | Write-Host -Fore Yellow
			$lines[$_i]        | Write-Host -Fore Green
			$lines[($_i+1)..$lines.Count] | Write-Host -Fore Yellow
			$_i-- # back a line, we're gonna step forward when we loop
		}
		"t" { # Time Check
			 $dur = [DateTime]::Now - $_StartTime
       Write-Host -Fore $promptColor $(
          "{3} -- $(if($dur.Hours -gt 0){'{0}h '})$(if($dur.Minutes -gt 0){'{1}m '}){2}s" -f
          $dur.Hours, $dur.Minutes, $dur.Seconds, ([DateTime]::Now.ToShortTimeString()))
			 $_i-- # back a line, we're gonna step forward when we loop
		}
		"s" { # Suspend (Enter Nested Prompt)
			Write-Host -Fore $promptColor "<Suspending demo - type 'Exit' to resume>"
			$Host.EnterNestedPrompt()
			$_i-- # back a line, we're gonna step forward when we loop
		}
		"g" { # GoTo Line Number
			$i = [int](Read-Host "line number")
			if($i -le $_lines.Count) {
				if($i -gt 0) {
               # extra line back because we're gonna step forward when we loop
               $_i = Rewind -lines $_lines -index $_i -steps (($_i-$i)+1)
				} else {
					$_i = -1 # Start negative, because we step forward when we loop
				}
			}
		}
		"f" { # Find by pattern
			$match = $_lines | Select-String (Read-Host "search string")
			if($null -eq $match) {
				Write-Host -Fore Red "Can't find a matching line"
			} else {
				$match | ForEach-Object { Write-Host -Fore $promptColor $("[{0,2}] {1}" -f ($_.LineNumber - 1), $_.Line) }
				if($match.Count -lt 1) {
					$_i = $match.lineNumber - 2  # back a line, we're gonna step forward when we loop
				} else {
					$_i-- # back a line, we're gonna step forward when we loop
				}
			}
		}
      "c" {
         Clear-Host
         $_i-- # back a line, we're gonna step forward when we loop
      }
		"$([char]13)" { # on enter
			Write-Host
			trap [System.Exception] {Write-Error $_; continue;}
			Invoke-Expression ($_lines[$_i]) | out-default
			if(-not $NoPauseAfterExecute -and -not $FullAuto) {
				$null = $RawUI.ReadKey("NoEcho,IncludeKeyUp")  # Pause after output for no apparent reason... ;)
			}
		}
		default
		{
			Write-Host -Fore Green "`nKey not recognized.  Press ? for help, or ENTER to execute the command."
			$_i-- # back a line, we're gonna step forward when we loop
		}
	}
}
$dur = [DateTime]::Now - $_StartTime
Write-Host -Fore $promptColor $(
   "<Demo Complete -- $(if($dur.Hours -gt 0){'{0}h '})$(if($dur.Minutes -gt 0){'{1}m '}){2}s>" -f
   $dur.Hours, $dur.Minutes, $dur.Seconds, [DateTime]::Now.ToLongTimeString())
Write-Host -Fore $promptColor $([DateTime]::now)
Write-Host
# SIG # Begin signature block
# MIIuzQYJKoZIhvcNAQcCoIIuvjCCLroCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAf6x7vgdYzRS1j
# w2ClPpGsdAdONe8P34xmsGt3o4irAaCCE24wggVyMIIDWqADAgECAhB2U/6sdUZI
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
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBqfmlk8C+bP8EkeqJCHpTvx31frcXg
# QI9bffqaKvbUSDA4BgorBgEEAYI3AgEMMSowKKACgAChIoAgaHR0cHM6Ly93d3cu
# bWlsZXN0b25lcHN0b29scy5jb20wDQYJKoZIhvcNAQEBBQAEggIAA6MiHNnKVtqP
# PLMHgZZ5Ddoa7r/1aCX/2aJbK6naz16Oe7f6NIdnkSrUOCY73onB7He2ybIH5lDe
# YbcSaSR1FQkkTSpGZuNFYrDG+A2c3CixY05QMoQZ0yTdGwHVm+Jb/SCHzWpImKu6
# OzbU+VDSVbeabuVyzpeiYF8fKwt7HRAoznFKUtbQ1Z9bVARo1VD6wW5ueeczGgc3
# IFGnIZpNW2htNziW0J2J6C0gvN70/N/K+tc1WPMbbtbrtNwz3B1EkqO7W69QxZlu
# onA8qH2Jgb8ZcKx+K02FKq3CTs3P3qTEJJzmETiJatduCudiXxl9g/6cQFktMXO1
# 1DhEXYYxhcQSUcGvx66SZNr/HOcII2g7ATxllBNCfiXgWCHWDPo5bkty9BQQHylK
# tFDn49h46YAy46xqKOfLIp620ar+EiBxTv4V6cJRSms+pPhGxJqFp9iQljP316f6
# 7vgC9U6IFm+ytcQYY+dnLTfYvzSLVFvIRhhYD4Q2unu0EOs6EM++ZK5VGtuRZvmf
# uvw7WZAToWQbao67yUQq4MRSTsjoAXJ5ZaGFg7x8PUqQWzR57K9N+D9S2jStOtOJ
# Dv6i2ivnaNoEQMTVxPgyyYAcKZx/fVG9KZQ+7EPdzRwjWdHrFpaHVq/M6YsBjPAj
# GWRGCAz2uUqTmROy3qwFzZHlvtn5Qhmhghd2MIIXcgYKKwYBBAGCNwMDATGCF2Iw
# ghdeBgkqhkiG9w0BBwKgghdPMIIXSwIBAzEPMA0GCWCGSAFlAwQCAQUAMHcGCyqG
# SIb3DQEJEAEEoGgEZjBkAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# xsTDH7+T941Szh9kLHhOu9BlAuKeSuvWCgOnJBvbC0MCEDgRXsERwhFFcOB8nQF5
# QrsYDzIwMjUwNjE4MTg1NzQ0WqCCEzowggbtMIIE1aADAgECAhAKgO8YS43xBYLR
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
# hvcNAQkFMQ8XDTI1MDYxODE4NTc0NFowKwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU
# 3WIwrIYKLTBr2jixaHlSMAf7QX4wLwYJKoZIhvcNAQkEMSIEIN0ckDbxXva1ubfc
# 3X8VT/hwrW1S81qzd1Xt4qv5ut+GMDcGCyqGSIb3DQEJEAIvMSgwJjAkMCIEIEqg
# P6Is11yExVyTj4KOZ2ucrsqzP+NtJpqjNPFGEQozMA0GCSqGSIb3DQEBAQUABIIC
# AMxUTLszqPGhb8O2ZwK0HToMCorUX6XKlBt2EQu4SnEP0HlWLunysp0QqIOE+hJh
# +47i4UJ20mwvTUREZ4wvRSb2YrYozBiyELTBq1CC8AqYv0qdTOuuyb78Qq2grxwc
# cjDlcxGGovzemh156ZuQuGaR1TkQjSkfrL0EISBr4WDukhiwLXh4fhT13lDUYmDt
# Tu39K0re5enCzo0J/Wqe8ZOZ+56b3vkE9h+KL2Za10Hd00mTmGRRxvgz2FyK9yuc
# hYS+oMvxGXdxJID09++HVrxqS3WK/nhC7KT8b7xguT+aoMYot29/bBGapBAS9KhM
# dSENPgBzwBHvkh+c3Lzilp84PNxoTQrdtha2FUOdz8rPDnQt5Vqz2KJulBEvjfU/
# /HLWJMDouqf1IHR8zRLHQ4qfZMrffIA+CplYa4c2mR6661COIx4rcvPlUYbt1N06
# zg9i4/fHp7dD7l5e+u/ocSaB9dsx/FhH0MbgYuwqrK/m+Oua+Y4Px1bEjOgK+qVE
# RM/0Zrm94ZTSwTGVsu9biYh894AvO/YmLSQ8Ai4F1O72mLvY4jZn5TfwxOWS9mt1
# PUbibPvSjP8Jp49XYwdWO9j7a3xxom5qlPd8Ho3MmqgN/DCld44hh0vyL3YHyg8j
# wix8dBDXqum5rF/cAsMgu1+h9be2pqFcWl8pQm/Akqtq
# SIG # End signature block
