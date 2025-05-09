' ParseWiki.vbs
' VBScript program to parse TechNet Wiki Article HTML for heading
' and anchor name tag problems.
' Script Author: Richard L. Mueller
' Version 1.0 - October 11, 2012
' Version 1.1 - November 8, 2012 - Improve recognition of headings.
' Version 1.2 - February 15, 2012 - Find foreign characters. Find: � � � �
' Version 1.3 - April 10, 2013 - Find rgb and hex color values.
' Version 1.4 - June 7, 2013 - Suggest best color to replace rgb color values.
' Version 1.5 - December 2, 2013 - Find any extended ASCII characters.

Option Explicit

Dim strFile, objFSO, objFile, strText, objRE, objHeadings, objHeading
Dim strHeading, objList, objRE2, objRE3, objRE4, objNames, objName, strName
Dim strLeading, objTags, objTag, strTag, intCount, strHeading2
Dim k, blnName, blnSpan, blnEndName, blnEndSpan, blnBlank, blnHeading
Dim intIndex, strTemp, objRGBColors, objRGBColor, objHexColors, objHexColor
Dim strColor, arrColors(138, 3)

Const ForReading = 1

' Check for file or prompt.
' This file should be a copy of the Wiki article HTML.
If (Wscript.Arguments.Count = 0) Then
    strFile = InputBox("Enter file name")
    ' Check if user canceled.
    If (strFile = "") Then
        Wscript.Quit
    End If
Else
    strFile = Wscript.Arguments(0)
End If

' Open the file.
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set objFile = objFSO.OpenTextFile(strFile, ForReading)
If (Err.Number <> 0) Then
    Wscript.Echo "Unable to open file: " & strFile
    Wscript.Echo "Error Number: " & Err.Number
    Wscript.Echo "Description: " & Err.Description
    Wscript.Quit
End If
On Error GoTo 0

' Populate array with the standard color names and their decimal RGB values.
arrColors(0, 0) = "Red"
arrColors(0, 1) = "255"
arrColors(0, 2) = "0"
arrColors(0, 3) = "0"

arrColors(1, 0) = "DarkRed"
arrColors(1, 1) = "139"
arrColors(1, 2) = "0"
arrColors(1, 3) = "0"

arrColors(2, 0) = "FireBrick"
arrColors(2, 1) = "178"
arrColors(2, 2) = "34"
arrColors(2, 3) = "34"

arrColors(3, 0) = "Crimson"
arrColors(3, 1) = "220"
arrColors(3, 2) = "20"
arrColors(3, 3) = "60"

arrColors(4, 0) = "IndianRed"
arrColors(4, 1) = "205"
arrColors(4, 2) = "92"
arrColors(4, 3) = "92"

arrColors(5, 0) = "LightCoral"
arrColors(5, 1) = "240"
arrColors(5, 2) = "128"
arrColors(5, 3) = "128"

arrColors(6, 0) = "DarkSalmon"
arrColors(6, 1) = "233"
arrColors(6, 2) = "150"
arrColors(6, 3) = "122"

arrColors(7, 0) = "Salmon"
arrColors(7, 1) = "250"
arrColors(7, 2) = "128"
arrColors(7, 3) = "114"

arrColors(8, 0) = "LightSalmon"
arrColors(8, 1) = "255"
arrColors(8, 2) = "160"
arrColors(8, 3) = "122"

arrColors(9, 0) = "MediumVioletRed"
arrColors(9, 1) = "199"
arrColors(9, 2) = "21"
arrColors(9, 3) = "133"

arrColors(10, 0) = "PaleVioletRed"
arrColors(10, 1) = "219"
arrColors(10, 2) = "112"
arrColors(10, 3) = "147"

arrColors(11, 0) = "DeepPink"
arrColors(11, 1) = "255"
arrColors(11, 2) = "20"
arrColors(11, 3) = "147"

arrColors(12, 0) = "HotPink"
arrColors(12, 1) = "255"
arrColors(12, 2) = "105"
arrColors(12, 3) = "180"

arrColors(13, 0) = "LightPink"
arrColors(13, 1) = "255"
arrColors(13, 2) = "182"
arrColors(13, 3) = "193"

arrColors(14, 0) = "Pink"
arrColors(14, 1) = "255"
arrColors(14, 2) = "192"
arrColors(14, 3) = "203"

arrColors(15, 0) = "OrangeRed"
arrColors(15, 1) = "255"
arrColors(15, 2) = "69"
arrColors(15, 3) = "0"

arrColors(16, 0) = "Tomato"
arrColors(16, 1) = "255"
arrColors(16, 2) = "99"
arrColors(16, 3) = "71"

arrColors(17, 0) = "Coral"
arrColors(17, 1) = "255"
arrColors(17, 2) = "127"
arrColors(17, 3) = "80"

arrColors(18, 0) = "DarkOrange"
arrColors(18, 1) = "255"
arrColors(18, 2) = "140"
arrColors(18, 3) = "0"

arrColors(19, 0) = "Orange"
arrColors(19, 1) = "255"
arrColors(19, 2) = "165"
arrColors(19, 3) = "0"

arrColors(20, 0) = "Gold"
arrColors(20, 1) = "255"
arrColors(20, 2) = "215"
arrColors(20, 3) = "0"

arrColors(21, 0) = "Yellow"
arrColors(21, 1) = "255"
arrColors(21, 2) = "255"
arrColors(21, 3) = "0"

arrColors(22, 0) = "LightYellow"
arrColors(22, 1) = "255"
arrColors(22, 2) = "255"
arrColors(22, 3) = "224"

arrColors(23, 0) = "LemonChiffon"
arrColors(23, 1) = "255"
arrColors(23, 2) = "250"
arrColors(23, 3) = "205"

arrColors(24, 0) = "LightGoldenrodYellow"
arrColors(24, 1) = "250"
arrColors(24, 2) = "250"
arrColors(24, 3) = "210"

arrColors(25, 0) = "PapayaWhip"
arrColors(25, 1) = "255"
arrColors(25, 2) = "239"
arrColors(25, 3) = "213"

arrColors(26, 0) = "Moccasin"
arrColors(26, 1) = "255"
arrColors(26, 2) = "228"
arrColors(26, 3) = "181"

arrColors(27, 0) = "PeachPuff"
arrColors(27, 1) = "255"
arrColors(27, 2) = "218"
arrColors(27, 3) = "185"

arrColors(28, 0) = "PaleGoldenrod"
arrColors(28, 1) = "238"
arrColors(28, 2) = "232"
arrColors(28, 3) = "170"

arrColors(29, 0) = "Khaki"
arrColors(29, 1) = "240"
arrColors(29, 2) = "230"
arrColors(29, 3) = "140"

arrColors(30, 0) = "DarkKhaki"
arrColors(30, 1) = "189"
arrColors(30, 2) = "183"
arrColors(30, 3) = "107"

arrColors(31, 0) = "CornSilk"
arrColors(31, 1) = "255"
arrColors(31, 2) = "248"
arrColors(31, 3) = "220"

arrColors(32, 0) = "BlanchedAlmond"
arrColors(32, 1) = "255"
arrColors(32, 2) = "235"
arrColors(32, 3) = "205"

arrColors(33, 0) = "Bisque"
arrColors(33, 1) = "255"
arrColors(33, 2) = "228"
arrColors(33, 3) = "196"

arrColors(34, 0) = "NavajoWhite"
arrColors(34, 1) = "255"
arrColors(34, 2) = "222"
arrColors(34, 3) = "173"

arrColors(35, 0) = "Wheat"
arrColors(35, 1) = "245"
arrColors(35, 2) = "222"
arrColors(35, 3) = "179"

arrColors(36, 0) = "BurlyWood"
arrColors(36, 1) = "222"
arrColors(36, 2) = "184"
arrColors(36, 3) = "135"

arrColors(37, 0) = "Tan"
arrColors(37, 1) = "210"
arrColors(37, 2) = "180"
arrColors(37, 3) = "140"

arrColors(38, 0) = "RosyBrown"
arrColors(38, 1) = "188"
arrColors(38, 2) = "143"
arrColors(38, 3) = "143"

arrColors(39, 0) = "SandyBrown"
arrColors(39, 1) = "244"
arrColors(39, 2) = "164"
arrColors(39, 3) = "96"

arrColors(40, 0) = "Goldenrod"
arrColors(40, 1) = "218"
arrColors(40, 2) = "165"
arrColors(40, 3) = "32"

arrColors(41, 0) = "DarkGoldenrod"
arrColors(41, 1) = "184"
arrColors(41, 2) = "134"
arrColors(41, 3) = "11"

arrColors(42, 0) = "Peru"
arrColors(42, 1) = "205"
arrColors(42, 2) = "133"
arrColors(42, 3) = "63"

arrColors(43, 0) = "Chocolate"
arrColors(43, 1) = "210"
arrColors(43, 2) = "105"
arrColors(43, 3) = "30"

arrColors(44, 0) = "SaddleBrown"
arrColors(44, 1) = "139"
arrColors(44, 2) = "69"
arrColors(44, 3) = "19"

arrColors(45, 0) = "Sienna"
arrColors(45, 1) = "160"
arrColors(45, 2) = "82"
arrColors(45, 3) = "45"

arrColors(46, 0) = "Brown"
arrColors(46, 1) = "165"
arrColors(46, 2) = "42"
arrColors(46, 3) = "42"

arrColors(47, 0) = "Maroon"
arrColors(47, 1) = "128"
arrColors(47, 2) = "0"
arrColors(47, 3) = "0"

arrColors(48, 0) = "DarkOliveGreen"
arrColors(48, 1) = "85"
arrColors(48, 2) = "107"
arrColors(48, 3) = "47"

arrColors(49, 0) = "Olive"
arrColors(49, 1) = "128"
arrColors(49, 2) = "128"
arrColors(49, 3) = "0"

arrColors(50, 0) = "OliveDrab"
arrColors(50, 1) = "107"
arrColors(50, 2) = "142"
arrColors(50, 3) = "35"

arrColors(51, 0) = "YellowGreen"
arrColors(51, 1) = "154"
arrColors(51, 2) = "205"
arrColors(51, 3) = "50"

arrColors(52, 0) = "LimeGreen"
arrColors(52, 1) = "50"
arrColors(52, 2) = "205"
arrColors(52, 3) = "50"

arrColors(53, 0) = "Lime"
arrColors(53, 1) = "0"
arrColors(53, 2) = "255"
arrColors(53, 3) = "0"

arrColors(54, 0) = "LawnGreen"
arrColors(54, 1) = "124"
arrColors(54, 2) = "252"
arrColors(54, 3) = "0"

arrColors(55, 0) = "Chartreuse"
arrColors(55, 1) = "127"
arrColors(55, 2) = "255"
arrColors(55, 3) = "0"

arrColors(56, 0) = "GreenYellow"
arrColors(56, 1) = "173"
arrColors(56, 2) = "255"
arrColors(56, 3) = "47"

arrColors(57, 0) = "SpringGreen"
arrColors(57, 1) = "0"
arrColors(57, 2) = "255"
arrColors(57, 3) = "127"

arrColors(58, 0) = "MediumSpringGreen"
arrColors(58, 1) = "0"
arrColors(58, 2) = "250"
arrColors(58, 3) = "154"

arrColors(59, 0) = "LightGreen"
arrColors(59, 1) = "144"
arrColors(59, 2) = "238"
arrColors(59, 3) = "144"

arrColors(60, 0) = "PaleGreen"
arrColors(60, 1) = "152"
arrColors(60, 2) = "251"
arrColors(60, 3) = "152"

arrColors(61, 0) = "DarkSeaGreen"
arrColors(61, 1) = "143"
arrColors(61, 2) = "188"
arrColors(61, 3) = "143"

arrColors(62, 0) = "MediumSeaGreen"
arrColors(62, 1) = "60"
arrColors(62, 2) = "179"
arrColors(62, 3) = "113"

arrColors(63, 0) = "SeaGreen"
arrColors(63, 1) = "46"
arrColors(63, 2) = "139"
arrColors(63, 3) = "87"

arrColors(64, 0) = "ForestGreen"
arrColors(64, 1) = "34"
arrColors(64, 2) = "139"
arrColors(64, 3) = "34"

arrColors(65, 0) = "Green"
arrColors(65, 1) = "0"
arrColors(65, 2) = "128"
arrColors(65, 3) = "0"

arrColors(66, 0) = "DarkGreen"
arrColors(66, 1) = "0"
arrColors(66, 2) = "100"
arrColors(66, 3) = "0"

arrColors(67, 0) = "MediumAquamarine"
arrColors(67, 1) = "102"
arrColors(67, 2) = "205"
arrColors(67, 3) = "170"

arrColors(68, 0) = "Aqua"
arrColors(68, 1) = "0"
arrColors(68, 2) = "255"
arrColors(68, 3) = "255"

arrColors(69, 0) = "Cyan"
arrColors(69, 1) = "0"
arrColors(69, 2) = "255"
arrColors(69, 3) = "255"

arrColors(70, 0) = "LightCyan"
arrColors(70, 1) = "224"
arrColors(70, 2) = "255"
arrColors(70, 3) = "255"

arrColors(71, 0) = "PaleTurquoise"
arrColors(71, 1) = "175"
arrColors(71, 2) = "238"
arrColors(71, 3) = "238"

arrColors(72, 0) = "Aquamarine"
arrColors(72, 1) = "127"
arrColors(72, 2) = "255"
arrColors(72, 3) = "212"

arrColors(73, 0) = "Turquoise"
arrColors(73, 1) = "64"
arrColors(73, 2) = "224"
arrColors(73, 3) = "208"

arrColors(74, 0) = "MediumTurquoise"
arrColors(74, 1) = "72"
arrColors(74, 2) = "209"
arrColors(74, 3) = "204"

arrColors(75, 0) = "DarkTurquoise"
arrColors(75, 1) = "0"
arrColors(75, 2) = "206"
arrColors(75, 3) = "209"

arrColors(76, 0) = "LightSeaGreen"
arrColors(76, 1) = "32"
arrColors(76, 2) = "178"
arrColors(76, 3) = "170"

arrColors(77, 0) = "CadetBlue"
arrColors(77, 1) = "95"
arrColors(77, 2) = "158"
arrColors(77, 3) = "160"

arrColors(78, 0) = "DarkCyan"
arrColors(78, 1) = "0"
arrColors(78, 2) = "139"
arrColors(78, 3) = "139"

arrColors(79, 0) = "Teal"
arrColors(79, 1) = "0"
arrColors(79, 2) = "128"
arrColors(79, 3) = "128"

arrColors(80, 0) = "LightSteelBlue"
arrColors(80, 1) = "176"
arrColors(80, 2) = "196"
arrColors(80, 3) = "222"

arrColors(81, 0) = "PowderBlue"
arrColors(81, 1) = "176"
arrColors(81, 2) = "224"
arrColors(81, 3) = "230"

arrColors(82, 0) = "LightBlue"
arrColors(82, 1) = "173"
arrColors(82, 2) = "216"
arrColors(82, 3) = "230"

arrColors(83, 0) = "SkyBlue"
arrColors(83, 1) = "135"
arrColors(83, 2) = "206"
arrColors(83, 3) = "235"

arrColors(84, 0) = "LightSkyBlue"
arrColors(84, 1) = "135"
arrColors(84, 2) = "206"
arrColors(84, 3) = "250"

arrColors(85, 0) = "DeepSkyBlue"
arrColors(85, 1) = "0"
arrColors(85, 2) = "191"
arrColors(85, 3) = "255"

arrColors(86, 0) = "DodgerBlue"
arrColors(86, 1) = "30"
arrColors(86, 2) = "144"
arrColors(86, 3) = "255"

arrColors(87, 0) = "CornflowerBlue"
arrColors(87, 1) = "100"
arrColors(87, 2) = "149"
arrColors(87, 3) = "237"

arrColors(88, 0) = "SteelBlue"
arrColors(88, 1) = "70"
arrColors(88, 2) = "130"
arrColors(88, 3) = "180"

arrColors(89, 0) = "RoyalBlue"
arrColors(89, 1) = "65"
arrColors(89, 2) = "105"
arrColors(89, 3) = "225"

arrColors(90, 0) = "Blue"
arrColors(90, 1) = "0"
arrColors(90, 2) = "0"
arrColors(90, 3) = "255"

arrColors(91, 0) = "MediumBlue"
arrColors(91, 1) = "0"
arrColors(91, 2) = "0"
arrColors(91, 3) = "205"

arrColors(92, 0) = "DarkBlue"
arrColors(92, 1) = "0"
arrColors(92, 2) = "0"
arrColors(92, 3) = "139"

arrColors(93, 0) = "Navy"
arrColors(93, 1) = "0"
arrColors(93, 2) = "0"
arrColors(93, 3) = "128"

arrColors(94, 0) = "MidnightBlue"
arrColors(94, 1) = "25"
arrColors(94, 2) = "25"
arrColors(94, 3) = "112"

arrColors(95, 0) = "Lavender"
arrColors(95, 1) = "230"
arrColors(95, 2) = "230"
arrColors(95, 3) = "250"

arrColors(96, 0) = "Thistle"
arrColors(96, 1) = "216"
arrColors(96, 2) = "191"
arrColors(96, 3) = "216"

arrColors(97, 0) = "Plum"
arrColors(97, 1) = "221"
arrColors(97, 2) = "160"
arrColors(97, 3) = "221"

arrColors(98, 0) = "Violet"
arrColors(98, 1) = "238"
arrColors(98, 2) = "130"
arrColors(98, 3) = "238"

arrColors(99, 0) = "Orchid"
arrColors(99, 1) = "218"
arrColors(99, 2) = "112"
arrColors(99, 3) = "214"

' Also called Fuchsia.
arrColors(100, 0) = "Magenta"
arrColors(100, 1) = "255"
arrColors(100, 2) = "0"
arrColors(100, 3) = "255"

arrColors(101, 0) = "MediumOrchid"
arrColors(101, 1) = "186"
arrColors(101, 2) = "85"
arrColors(101, 3) = "211"

arrColors(102, 0) = "MediumPurple"
arrColors(102, 1) = "147"
arrColors(102, 2) = "112"
arrColors(102, 3) = "219"

arrColors(103, 0) = "BlueViolet"
arrColors(103, 1) = "138"
arrColors(103, 2) = "43"
arrColors(103, 3) = "226"

arrColors(104, 0) = "DarkViolet"
arrColors(104, 1) = "148"
arrColors(104, 2) = "0"
arrColors(104, 3) = "211"

arrColors(105, 0) = "DarkOrchid"
arrColors(105, 1) = "153"
arrColors(105, 2) = "50"
arrColors(105, 3) = "204"

arrColors(106, 0) = "DarkMagenta"
arrColors(106, 1) = "139"
arrColors(106, 2) = "0"
arrColors(106, 3) = "139"

arrColors(107, 0) = "Purple"
arrColors(107, 1) = "128"
arrColors(107, 2) = "0"
arrColors(107, 3) = "128"

arrColors(108, 0) = "Indigo"
arrColors(108, 1) = "75"
arrColors(108, 2) = "0"
arrColors(108, 3) = "130"

arrColors(109, 0) = "DarkSlateBlue"
arrColors(109, 1) = "72"
arrColors(109, 2) = "61"
arrColors(109, 3) = "139"

arrColors(110, 0) = "SlateBlue"
arrColors(110, 1) = "106"
arrColors(110, 2) = "90"
arrColors(110, 3) = "205"

arrColors(111, 0) = "MediumSlateBlue"
arrColors(111, 1) = "123"
arrColors(111, 2) = "104"
arrColors(111, 3) = "238"

arrColors(112, 0) = "White"
arrColors(112, 1) = "255"
arrColors(112, 2) = "255"
arrColors(112, 3) = "255"

arrColors(113, 0) = "Snow"
arrColors(113, 1) = "255"
arrColors(113, 2) = "250"
arrColors(113, 3) = "250"

arrColors(114, 0) = "HoneyDew"
arrColors(114, 1) = "240"
arrColors(114, 2) = "255"
arrColors(114, 3) = "240"

arrColors(115, 0) = "MintCream"
arrColors(115, 1) = "245"
arrColors(115, 2) = "255"
arrColors(115, 3) = "250"

arrColors(116, 0) = "Azure"
arrColors(116, 1) = "240"
arrColors(116, 2) = "255"
arrColors(116, 3) = "255"

arrColors(117, 0) = "AliceBlue"
arrColors(117, 1) = "240"
arrColors(117, 2) = "248"
arrColors(117, 3) = "255"

arrColors(118, 0) = "GhostWhite"
arrColors(118, 1) = "248"
arrColors(118, 2) = "248"
arrColors(118, 3) = "255"

arrColors(119, 0) = "WhiteSmoke"
arrColors(119, 1) = "245"
arrColors(119, 2) = "245"
arrColors(119, 3) = "245"

arrColors(120, 0) = "Seashell"
arrColors(120, 1) = "255"
arrColors(120, 2) = "245"
arrColors(120, 3) = "238"

arrColors(121, 0) = "Beige"
arrColors(121, 1) = "245"
arrColors(121, 2) = "245"
arrColors(121, 3) = "220"

arrColors(122, 0) = "OldLace"
arrColors(122, 1) = "253"
arrColors(122, 2) = "245"
arrColors(122, 3) = "230"

arrColors(123, 0) = "FloralWhite"
arrColors(123, 1) = "255"
arrColors(123, 2) = "250"
arrColors(123, 3) = "240"

arrColors(124, 0) = "Ivory"
arrColors(124, 1) = "255"
arrColors(124, 2) = "255"
arrColors(124, 3) = "240"

arrColors(125, 0) = "AntiqueWhite"
arrColors(125, 1) = "250"
arrColors(125, 2) = "235"
arrColors(125, 3) = "215"

arrColors(126, 0) = "Linen"
arrColors(126, 1) = "250"
arrColors(126, 2) = "240"
arrColors(126, 3) = "230"

arrColors(127, 0) = "LavenderBlush"
arrColors(127, 1) = "255"
arrColors(127, 2) = "240"
arrColors(127, 3) = "245"

arrColors(128, 0) = "MistyRose"
arrColors(128, 1) = "255"
arrColors(128, 2) = "228"
arrColors(128, 3) = "225"

arrColors(129, 0) = "Gainsboro"
arrColors(129, 1) = "220"
arrColors(129, 2) = "220"
arrColors(129, 3) = "220"

arrColors(130, 0) = "LightGray"
arrColors(130, 1) = "211"
arrColors(130, 2) = "211"
arrColors(130, 3) = "211"

arrColors(131, 0) = "Silver"
arrColors(131, 1) = "192"
arrColors(131, 2) = "192"
arrColors(131, 3) = "192"

arrColors(132, 0) = "DarkGray"
arrColors(132, 1) = "169"
arrColors(132, 2) = "169"
arrColors(132, 3) = "169"

arrColors(133, 0) = "Gray"
arrColors(133, 1) = "128"
arrColors(133, 2) = "128"
arrColors(133, 3) = "128"

arrColors(134, 0) = "DimGray"
arrColors(134, 1) = "105"
arrColors(134, 2) = "105"
arrColors(134, 3) = "105"

arrColors(135, 0) = "LightSlateGray"
arrColors(135, 1) = "119"
arrColors(135, 2) = "136"
arrColors(135, 3) = "153"

arrColors(136, 0) = "SlateGray"
arrColors(136, 1) = "112"
arrColors(136, 2) = "128"
arrColors(136, 3) = "144"

arrColors(137, 0) = "DarkSlateGray"
arrColors(137, 1) = "47"
arrColors(137, 2) = "79"
arrColors(137, 3) = "79"

arrColors(138, 0) = "Black"
arrColors(138, 1) = "0"
arrColors(138, 2) = "0"
arrColors(138, 3) = "0"

' Set up dictionary object.
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare

' Read the file contents. Convert to all lower case.
strText = LCase(objFile.ReadAll)
objFile.Close

Set objRE = New RegExp
objRE.Global = True
Set objRE2 = New RegExp
objRE2.Global = True
Set objRE3 = New RegExp
objRE3.Global = True
Set objRE4 = New RegExp
objRE4.Global = True

' Parse headings.
objRE.Pattern = "<h[0-9].*</h[0-9]>"
Set objHeadings = objRE.Execute(strText)

' Parse anchor name tags.
objRE2.Pattern = "<a name=.*></a>"

' Parse rgb color values.
objRE3.Pattern = "rgb\([0-9]+, [0-9]+, [0-9]+\)"

' Parse hex color values.
objRE4.Pattern = "#[0-9a-f]{6}"

' Read file for headings and anchor name tags.
Wscript.Echo "----- Analyze Headings"
intCount = 0
For Each objHeading In objHeadings
    strHeading = objHeading.Value
    ' Find anchor name tags in the heading.
    Set objNames = objRE2.Execute(strHeading)
    For Each objName  In objNames
        strName = objName.Value
        ' Check leading character of the anchor name tag.
        strLeading = Mid(strName, 10, 1)
        If (IsNumeric(strLeading) = True) Then
            Wscript.Echo "## Leading digit in anchor name tag: " & strHeading
            intCount = intCount + 1
        End If
        ' Check for embedded "0" characters in the anchor name tag.
        If (InStr(strName, "0") > 0) Then
            ' Remove any duplicate <a name>...</a> anchor tags for readability.
            intIndex = InStr(strName, "</a>")
            strTemp = Left(strName, intIndex + 3)
            ' "0" characters are only a problem in the first <a name> tag.
            If (InStr(strTemp, "0") > 0) Then
                Wscript.Echo "## Embedded ""0"" in name tag: " _
                    & Replace(strHeading, strName, strTemp)
                intCount = intCount + 1
            End If
        End If
        If (Foreign(strName) = True) Then
            Wscript.Echo "## Foreign character in anchor name tag: " & strHeading
            intCount = intCount + 1
        End If
        ' Check for duplicate anchor name tags.
        If (objList.Exists(strName) = False) Then
            objList.Add strName, True
        Else
            Wscript.Echo "## Duplicate anchor name tag: " & strHeading
            intCount = intCount + 1
        End If
    Next
    ' Check for blank heading (no characters outside name or span tags).
    blnName = False
    blnSpan = False
    blnEndName = False
    blnEndSpan = False
    blnHeading = True
    blnBlank = True
    ' Ignore spaces when looking for blanks.
    strHeading2 = Replace(strHeading, "&nbsp;", "")
    For k = 1 to Len(strHeading)
        If (blnHeading = False) Then
            If (Mid(strHeading2, k, 3) = "</h") Then
                Exit For
            End If
            If (blnName = True) And (Mid(strHeading2, k, 1) = ">") Then
                blnName = False
            End If
            If (Mid(strHeading2, k, 8) = "<a name=") Then
                blnName = True
            End If
            If (blnSpan = True) And (Mid(strHeading2, k, 1) = ">") Then
                blnSpan = False
            End If
            If (Mid(strHeading2, k, 12) = "<span style=") Then
                blnSpan = True
            End If
            If (blnEndName = True) And (Mid(strHeading2, k, 1) = ">") Then
                blnEndName = False
            End If
            If (Mid(strHeading2, k, 4) = "</a>") Then
                blnEndName = True
            End If
            If (blnEndSpan = True) And (Mid(strHeading2, k, 1) = ">") Then
                blnEndSpan = False
            End If
            If (Mid(strHeading2, k, 7) = "</span>") Then
                blnEndSpan = True
            End If
            If (blnName = False) And (blnSpan = False) _
                    And (blnEndName = False) And (blnEndSpan = False) Then
                If (Mid(strHeading2, k, 1) <> ">") Then
                    blnBlank = False
                    Exit For
                End If
            End If
        End If
        If (blnHeading = True) And (Mid(strHeading2, k, 1) = ">") Then
            blnHeading = False
        End If
    Next
    If (blnBlank = True) Then
       Wscript.Echo "## Blank heading: " & strHeading
        intCount = intCount + 1
    End If
Next
Wscript.Echo "Number of problems in headings: " & CStr(intCount)

' Read all anchor name tags in the file.
Wscript.Echo "----- Analyze Anchor Name Tags"
intCount = 0
objList.RemoveAll
Set objTags = objRE2.Execute(strText)
For Each objTag In objTags
    strTag = objTag.Value
    ' Check for duplicates.
    If (objList.Exists(strTag) = False) Then
        objList.Add strTag, True
    Else
        Wscript.Echo "## Duplicate anchor name tag: " & strTag
        intCount = intCount + 1
    End If
    ' Check for embedded "0" characters in the anchor name tag.
    If (InStr(strTag, "0") > 0) Then
        ' Remove any duplicate <a name>...</a> anchor tags for readability.
        intIndex = InStr(strTag, "</a>")
        strTemp = Left(strTag, intIndex + 3)
        ' "0" characters are only a problem in the first <a name> tag.
        If (InStr(strTemp, "0") > 0) Then
            Wscript.Echo "## Embedded ""0"" in name tag: " & strTemp
            intCount = intCount + 1
        End If
    End If
    If (Foreign(strTag) = True) Then
        ' Remove any duplicate <a name>...</a> anchor tags for readability.
        intIndex = InStr(strTag, "</a>")
        strTemp = Left(strTag, intIndex + 3)
        Wscript.Echo "## Foreign character in anchor name tag: " & strTemp
        intCount = intCount + 1
    End If

Next
Wscript.Echo "Number of problems in anchor name tags: " & CStr(intCount)

Wscript.Echo "----- Search for color values"
objList.RemoveAll
intCount = 0

' To select digits from RGB values.
objRE.Pattern = "[0-9]+"

Set objRGBColors = objRE3.Execute(strText)
For Each objRGBColor In objRGBColors
    strColor = objRGBColor.Value
    ' Check for duplicates.
    If (objList.Exists(strColor) = False) Then
        objList.Add strColor, True
        Wscript.Echo "## Found rgb color: " & strColor _
            & " [best standard color: " & BestColor(strColor) & "]"
        intCount = intCount + 1
    End If
Next

objList.RemoveAll
Set objHexColors = objRE4.Execute(strText)
For Each objHexColor In objHexColors
    strColor = objHexColor.Value
    ' Check for duplicates.
    If (objList.Exists(strColor) = False) Then
        objList.Add strColor, True
        Wscript.Echo "## Found hex color value: " & UCase(strColor)
        intCount = intCount + 1
    End If
Next
Wscript.Echo "Total color values found: " & CStr(intCount)

Function BestColor(ByVal strRGBColor)
    ' Select best color for RGB values.
    ' The following must have global scope and be defined in the main
    ' program: objRE, arrColors

    Dim objMatches, objMatch, k, lngMin, lngValue, intColor, arrValues(2)

    Set objMatches = objRE.Execute(strRGBColor)
    k = 0
    For Each objMatch In objMatches
        arrValues(k) = objMatch.Value
        k = k + 1
    Next

    ' Consider all of the standard colors.
    lngMin = 3*255*255
    For k = 0 To 138
        ' Calculate the distance between this standard color and the RGB value.
        lngValue = (arrColors(k, 1) - arrValues(0))^2 _
            + (arrColors(k, 2) - arrValues(1))^2 _
            + (arrColors(k, 3) - arrValues(2))^2
        ' Keep track of the color with the smallest distance (least difference).
        If (lngValue <= lngMin) Then
            intColor = k
            lngMin = lngValue
        End If
    Next
    ' Select the best standard color name.
    BestColor = arrColors(intColor, 0)

End Function

Function Foreign(strValue)
    ' Function to detect foreign characters.
    ' Returns True if the input string contains any extended ASCII characters,
    ' otherwise returns False.
    Dim k

    Foreign = False
    For k = 1 To Len(strValue)
        If (Asc(Mid(strValue, k, 1)) > 127) Then
            Foreign = True
            Exit Function
        End If
    Next
End Function
