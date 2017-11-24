
' List of API support

' List of EXTs API support for all modes

'#Const EXT_IUTIL = 1
'#Const EXT_ICOMMAND = 1
'#Const EXT_ICINIFILE = 1
'#Const EXT_ITIMER = 1
'#Const EXT_HKTIMER = 1

' List of EXTs API support for mp mode only

'#Const EXT_IHALOENGINE = 1
'#Const EXT_IOBJECT = 1      ' Require EXT_IUTIL
'#Const EXT_IPLAYER = 1      ' Require EXT_IOBJECT; If define EXT_IADMIN, EXT_IPLAYER test will Not process
'#Const EXT_IADMIN = 1       ' Require EXT_IUTIL

' Not included in this UnitTest.
'#Const EXT_IDATABASE = 1
'#Const EXT_IDATABASESTATEMENT = 1
'#Const EXT_HKDATABASE = 1

' Future API support

'#Const EXT_INETWORK = 1            ' Will require mp mode test And possible client?
'#Const EXT_ISOUND = 1              ' Require client side test only.
'#Const EXT_IDIRECTX9 = 1           ' Require client side test only
'#Const EXT_HKEXTERNAL = 1          ' TBD

#If DO_NOT_INCLUDE_THIS Then
addon_info EXTPluginInfo = { "UnitTest Visual Basic", "1.0.0.0",
                            "DZS|All-In-One, founder of DZS",
                            "Used for verifying each API are working right in VB.NET language under C99 standard.",
                            "UnitTest",
                            "unit_test",
                            "test_unit",
                            "unit test",
                            "[unit]test",
                            "test[unit]"};
#End If


' * Verification list as of 0.5.3.3
' *
' * EXT_IHALOENGINE          - Passed (except a few functions are not included in test.)
' * EXT_IOBJECT              - Passed (except a few functions are not included in test.)
' * EXT_IPLAYER              - Passed (except a few functions are not included in test.)
' * EXT_IADMIN               - Passed
' * EXT_ICOMMAND             - Passed
' * EXT_IDATABASE            - Not included in this UnitTest.
' * EXT_IDATABASESTATEMENT   - Not included in this UnitTest.
' * EXT_HKDATABASE           - Not included in this UnitTest.
' * EXT_ICINIFILE            - Passed
' * EXT_ITIMER               - Passed (Expect imbalance tick synchronize for 1/30 ticks per second after first load time.)
' * EXT_HKTIMER              - Passed
' * EXT_IUTIL                - Passed (except a few functions are not included in test.)
' * Future API support
' * EXT_INETWORK
' * EXT_ISOUND
' * EXT_IDIRECTX9
' * EXT_HKEXTERNAL
' 


'
' * This link is for effective usage in unmanaged code (not for C# code) to load managed dll.
' * http://stackoverflow.com/questions/773476/how-to-split-dot-net-hosting-function-when-calling-via-c-dll
' 


Imports System.Text
Imports System.Windows.Forms

Imports RGiesecke.DllExport
Imports System.Runtime.InteropServices

Namespace UnitTestCSharp
    Public Class Addon
        Public Shared hash As UInteger
#If EXT_IUTIL Then
        'boolean test section
        Public Shared trueStr As String = "true"
        Public Shared TRUEStrUpper As String = "TRUE"
        Public Shared trUeStrRandom As String = "trUe"
        Public Shared trueNumStr As String = "1"
        Public Shared falseStr As String = "false"
        Public Shared FALSEStrUpper As String = "FALSE"
        Public Shared faLseStrRandom As String = "faLse"
        Public Shared falseNumStr As String = "0"

        'team test section
        Public Shared blueStr As String = "blue"
        Public Shared BLUEStrUpper As String = "BLUE"
        Public Shared btStr As String = "bt"
        Public Shared redStr As String = "red"
        Public Shared REDStrUpper As String = "RED"
        Public Shared rtStr As String = "rt"

        'detect string test section
        Public Shared lettersStr As String = "lEtteRs"
        Public Shared letters2Str As String = "LeTterS"
        Public Shared numbersStr As String = "12348765"
        Public Shared numbers2Str As String = "87651234"
        Public Shared floatStr As String = "1234.8765"
        Public Shared doubleStr As String = "1.2348765"
        Public Shared hashStr As String = "87a651234"
        Public Shared hash2Str As String = "876512z34"

        Public Shared MatterStr As String = "Matter"
        Public Shared MattarStr As String = "Mattar"
        Public Shared MattarReplaceBeforeStr As String = "MattarTest 'Foobar'"
        Public Shared MatterReplaceBeforeStr As String = "MatterTest 'Foobar'"

        'directory and file test section
        Public Shared dirExtension As String = "extension"
        Public Shared fileHExt As String = "H-Ext.dll"
        Public Shared dirExtesion As String = "extesion"
        Public Shared fileHEt As String = "H-Et.dll"

        Public Shared replaceTestStr As New StringBuilder("Test 'Foobar'")
        Public Shared replaceBeforeStr As String = "Test 'Foobar'"
        Public Shared replaceAfterStr As String = "Test ''Foobar''"

        'regex test section
        Public Shared regexTestNoDB As New StringBuilder("? *? {test} )(string]here[there", 40)
        Public Shared regexTestNoDBAfter As String = ". .*. \{test\} \)\(string\]here\[there"
        Public Shared regexTestDB As New StringBuilder("? *? {test} )(string]here[there", 40)
        Public Shared regexTestDBAfter As String = "_ %_ \{test\} \)\(string\]here\[there"
        Public Shared wildcard As String = ".*"
        Public Shared wildcardEndTest As String = ".*Test"
        Public Shared wildcardBeginUnit As String = "Unit .*"
        Public Shared dotdotdot As String = "..."
        Public Shared hi_ As String = "Hi!"
        Public Shared Unit_TestUpper As String = "Unit Test"
        Public Shared unit_test As String = "unit test"

        'variant test section - CSharp only provide unicode input, absolutely no ansi support at all.
        'public static string variantFormatExpected = "Aa 1.000000 2.000002 1 25 50 4294967295 2147483647 2147483647 4294967295 aA";
        'public static string variantFormat = "{0:s} {2:f} {3:f} {4:hhd} {5:hd} {6:hu} {8:u} {7:d} {9:ld} {10:lu} {1:s}";
        Public Shared variantFormatExpected As String = "1.000000 2.000002 1 25 50 4294967295 2147483647 2147483647 4294967295 aA"
        Public Shared variantFormat As String = "{2:f} {3:f} {4:hhd} {5:hd} {6:hu} {8:u} {7:d} {9:ld} {10:lu} {1:s}"
#End If
#If EXT_ICOMMAND OrElse EXT_ICINIFILE Then
        Public Shared sectors As New Addon_API.addon_section_names() With {
            .sect_name1 = "unit_test",
            .sect_name2 = "test_unit",
            .sect_name3 = "unit test",
            .sect_name4 = "[unit]test",
            .sect_name5 = "test[unit]"
        }
#End If
#If EXT_ICINIFILE Then
        Public Shared iniFileStr As String = "UnitTestC.ini"
        Public Shared firstUnitTestCStr As String = "First Unit Test C"
        Public Shared str1_0 As String = "1.0"
        Public Shared str1_1 As String = "1.1"
        Public Shared str1_2 As String = "1.2"
        Public Shared str1_3 As String = "1.3"
        Public Shared iniFileDataStr As String = " [unit_test]" & vbCr & vbLf & " 1.0=First Unit Test C" & vbCr & vbLf & " [test_unit]" & vbCr & vbLf & " 1.1=First Unit Test C" & vbCr & vbLf & " [unit test]"
#End If

        'ICommand test section
#If EXT_ICOMMAND Then
        Public Shared eaoTestExecuteStr As String = "eao_test_execute"
        Public Shared eaoTestExecuteAliasStr As String = "testexec"
        Public Shared eaoLoadFileStr As String = "unit_test.txt"
        'This is needed in order to preserve function pointer address
        Public Shared eao_testExecutePtr As Addon_API.CmdFunc
        Public Shared eao_testExecuteOverridePtr As Addon_API.CmdFunc
        Public Shared eao_testExecuteOverride2Ptr As Addon_API.CmdFunc
        Public Shared Function eao_testExecute(<[In]> plI As Addon_API.PlayerInfo, <[In], Out> ByRef arg As Addon_API.ArgContainerVars, <[In]> protocolMsg As Addon_API.MSG_PROTOCOL, <[In]> idTimer As UInteger, <[In]> showChat As boolOption) As Addon_API.CMD_RETURN
            Return Addon_API.CMD_RETURN.SUCC
        End Function
        Public Shared Function eao_testExecuteOverride(<[In]> plI As Addon_API.PlayerInfo, <[In], Out> ByRef arg As Addon_API.ArgContainerVars, <[In]> protocolMsg As Addon_API.MSG_PROTOCOL, <[In]> idTimer As UInteger, <[In]> showChat As boolOption) As Addon_API.CMD_RETURN
            Return Addon_API.CMD_RETURN.SUCC
        End Function
        Public Shared Function eao_testExecuteOverride2(<[In]> plI As Addon_API.PlayerInfo, <[In], Out> ByRef arg As Addon_API.ArgContainerVars, <[In]> protocolMsg As Addon_API.MSG_PROTOCOL, <[In]> idTimer As UInteger, <[In]> showChat As boolOption) As Addon_API.CMD_RETURN
            Return Addon_API.CMD_RETURN.SUCC
        End Function
#End If
        'IPlayer test section
#If EXT_IPLAYER Then
        Public Shared cdHashKeyA As New StringBuilder(&H60)
#End If
        'IAdmin test section
#If EXT_IADMIN Then
        Public Shared username As String = "unittest"
        Public Shared usernamebad As String = "unittes"
        Public Shared cmdEaoLoad As String = "ext_addon_load unittest"
        Public Shared noKeyHere As String = "nokeyhere"
        Public Shared localhost As String = "127.0.0.1"
#End If
        'ITimer test section
#If EXT_ITIMER Then
        Public Shared pITimer As Addon_API.ITimer
        Public Shared TimerID As UInteger() = {0, 0, 0, 0}
        Public Shared TimerTickStart As UInteger = 0
        Public Shared TimerTickSys As UInteger() = {0, 0, 0, 0}
#End If
        'IHaloEngine test section
#If EXT_IHALOENGINE Then
        Public Shared rconTestStr As String = "Rcon Test"
        Public Shared playerChatTest As String = "Player Chat Test"
        Public Shared globalChatTest As String = "Global Chat Test"
        Public Shared password As String = "unitest"
        Public Shared passwordWGet As New StringBuilder("deadbeef", 8)
        Public Shared passwordAGet As New StringBuilder("deadbeef", 8)
#End If
        <DllExport("EXTOnEAOLoad", CallingConvention:=CallingConvention.Cdecl)>
        Public Shared Function EXTOnEAOLoad(uniqueHash As UInteger) As Addon_API.EAO_RETURN
            hash = uniqueHash
            Dim retCode As UInteger
            Dim testPtr As IntPtr = IntPtr.Zero
            Dim plI As New Addon_API.PlayerInfo(), plIKeep As New Addon_API.PlayerInfo(), plINull As New Addon_API.PlayerInfo()
#Region "IUtil test section"
#If EXT_IUTIL Then
            Dim pIUtil As Addon_API.IUtil = Addon_API.[Interface].getIUtil(hash)
            Try
                If pIUtil.isNotNull() Then
                    Addon_API.[Global].pIUtil = pIUtil
                    'm_allocMem & m_freeMem functions are not needed here.
                    Dim testBStrW As New StringBuilder(&H30)
                    Dim testBStrA As New StringBuilder(&H30)
                    Dim testStr1 As String = "Test String"
                    Dim matterStr__1 As New StringBuilder("Matter")
                    If pIUtil.m_strcatW(testBStrW, CUInt(testBStrW.Capacity), testStr1) <> 11 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpW(testBStrW.ToString(), testStr1) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strcatA(testBStrA, CUInt(testBStrA.Capacity), testStr1) <> 11 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpA(testBStrA.ToString(), testStr1) Then
                        Throw New ArgumentException()
                    End If
#Region "boolean values"
                    If pIUtil.m_strToBooleanW(testBStrW.ToString()) <> Addon_API.e_boolean.FAIL Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(testBStrA.ToString()) <> Addon_API.e_boolean.FAIL Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(trueStr) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(TRUEStrUpper) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(trueStr) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(trueNumStr) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(falseStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(falseStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(falseStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanW(falseNumStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(trueStr) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(TRUEStrUpper) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(trUeStrRandom) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(trueNumStr) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(falseStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(falseStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(faLseStrRandom) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToBooleanA(falseNumStr) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
#End Region
#Region "team values"
                    If pIUtil.m_strToTeamW(testBStrW.ToString()) <> e_color_team_index.TEAM_NONE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(testBStrA.ToString()) <> e_color_team_index.TEAM_NONE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamW(blueStr) <> e_color_team_index.TEAM_BLUE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamW(BLUEStrUpper) <> e_color_team_index.TEAM_BLUE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamW(btStr) <> e_color_team_index.TEAM_BLUE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamW(redStr) <> e_color_team_index.TEAM_RED Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamW(redStr) <> e_color_team_index.TEAM_RED Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamW(rtStr) <> e_color_team_index.TEAM_RED Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(blueStr) <> e_color_team_index.TEAM_BLUE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(BLUEStrUpper) <> e_color_team_index.TEAM_BLUE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(btStr) <> e_color_team_index.TEAM_BLUE Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(redStr) <> e_color_team_index.TEAM_RED Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(REDStrUpper) <> e_color_team_index.TEAM_RED Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strToTeamA(rtStr) <> e_color_team_index.TEAM_RED Then
                        Throw New ArgumentException()
                    End If
#End Region
#Region "Strings values verification"
                    testBStrA.Remove(0, testBStrA.Length)
                    pIUtil.m_toCharA(testBStrW.ToString(), testBStrW.Length + 1, testBStrA)
                    If Not pIUtil.m_strcmpA(testStr1, testBStrA.ToString()) Then
                        Throw New ArgumentException()
                    End If
                    pIUtil.m_toCharW(testBStrA.ToString(), testBStrA.Length + 1, testBStrW)
                    If Not pIUtil.m_strcmpW(testStr1, testBStrW.ToString()) Then
                        Throw New ArgumentException()
                    End If
                    testBStrA.Replace("t"c, "T"c)
                    testBStrW.Replace("t"c, "T"c)
                    If Not pIUtil.m_stricmpA(testStr1, testBStrA.ToString()) Then
                        Throw New ArgumentException()
                    End If
                    testBStrW.Remove(0, testBStrW.Length)
                    pIUtil.m_toCharW(testBStrA.ToString(), testBStrA.Length + 1, testBStrW)
                    If Not pIUtil.m_stricmpW(testStr1, testBStrW.ToString()) Then
                        Throw New ArgumentException()
                    End If
                    testBStrA.Replace("T"c, "t"c)
                    testBStrW.Replace("T"c, "t"c)
                    If Not pIUtil.m_stricmpW(lettersStr, letters2Str) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_stricmpA(lettersStr, letters2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strcmpW(lettersStr, letters2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strcmpA(lettersStr, letters2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_stricmpW(numbersStr, numbers2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_stricmpA(numbersStr, numbers2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strcmpW(numbersStr, numbers2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_strcmpA(numbersStr, numbers2Str) Then
                        Throw New ArgumentException()
                    End If

                    Dim [boolean] As Addon_API.e_boolean = pIUtil.m_shiftStrW(matterStr__1, 1, 3, 1, False)
                    If [boolean] <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_findSubStrFirstW(matterStr__1.ToString(), MatterStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_findSubStrFirstW(matterStr__1.ToString(), MattarStr) Then
                        Throw New ArgumentException()
                    End If
                    [boolean] = pIUtil.m_shiftStrW(matterStr__1, 1, 1, 3, True)
                    If [boolean] <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_findSubStrFirstW(matterStr__1.ToString(), MattarStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_findSubStrFirstW(matterStr__1.ToString(), MatterStr) Then
                        Throw New ArgumentException()
                    End If

                    'No reason to have 2 matter string since ANSII and Unicode are done by C# itself.
                    matterStr__1.Remove(0, matterStr__1.Length)
                    matterStr__1.Insert(0, MatterStr)
                    [boolean] = pIUtil.m_shiftStrA(matterStr__1, 1, 3, 1, False)
                    If [boolean] <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_findSubStrFirstA(matterStr__1.ToString(), MatterStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_findSubStrFirstA(matterStr__1.ToString(), MattarStr) Then
                        Throw New ArgumentException()
                    End If
                    [boolean] = pIUtil.m_shiftStrA(matterStr__1, 1, 1, 3, True)
                    If [boolean] <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_findSubStrFirstA(matterStr__1.ToString(), MattarStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_findSubStrFirstA(matterStr__1.ToString(), MatterStr) Then
                        Throw New ArgumentException()
                    End If

                    testBStrW.Remove(0, testBStrW.Length)
                    testBStrA.Remove(0, testBStrA.Length)
                    retCode = pIUtil.m_strcatW(testBStrW, 48, MattarStr)
                    If retCode <> 6 Then
                        Throw New ArgumentException()
                    End If
                    retCode = pIUtil.m_strcatA(testBStrA, 48, MatterStr)
                    If retCode <> 6 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpW(testBStrW.ToString(), MattarStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpA(testBStrA.ToString(), MatterStr) Then
                        Throw New ArgumentException()
                    End If
                    retCode = pIUtil.m_strcatW(testBStrW, 48, replaceBeforeStr)
                    If retCode <> 13 Then
                        Throw New ArgumentException()
                    End If
                    retCode = pIUtil.m_strcatA(testBStrA, 48, replaceBeforeStr)
                    If retCode <> 13 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpW(testBStrW.ToString(), MattarReplaceBeforeStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpA(testBStrA.ToString(), MatterReplaceBeforeStr) Then
                        Throw New ArgumentException()
                    End If
#End Region
#Region "isLetters, isFloat, isDouble, isNumbers, and isHash"
                    If Not pIUtil.m_isLettersW(lettersStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isLettersW(hashStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isLettersA(letters2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isLettersA(hash2Str) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIUtil.m_isNumberW(numbersStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isNumberW(hashStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isNumberW(floatStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isNumberA(numbers2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isNumberW(hash2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isNumberW(floatStr) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIUtil.m_isDoubleW(doubleStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isDoubleW(numbersStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isDoubleW(hashStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isDoubleW(floatStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isDoubleA(doubleStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isDoubleA(numbers2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isDoubleA(hash2Str) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isDoubleA(floatStr) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIUtil.m_isFloatW(floatStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isFloatW(doubleStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isFloatW(numbersStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isFloatW(hashStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isFloatA(floatStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isFloatA(doubleStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isFloatA(numbers2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isFloatA(hash2Str) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIUtil.m_isHashW(hashStr) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isHashW(floatStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isHashA(hash2Str) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isHashA(floatStr) Then
                        Throw New ArgumentException()
                    End If
#End Region
#Region "file & directory check"
                    If Not pIUtil.m_isDirExist(dirExtension, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If retCode > 0 Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isDirExist(dirExtesion, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isDirExist(fileHExt, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_isFileExist(fileHExt, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If retCode > 0 Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isFileExist(fileHEt, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_isFileExist(dirExtension, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
#End Region
#Region "Replace & undo relative + database regex replace."
                    pIUtil.m_replaceW(replaceTestStr)
                    If Not pIUtil.m_strcmpW(replaceTestStr.ToString(), replaceAfterStr) Then
                        Throw New ArgumentException()
                    End If
                    pIUtil.m_replaceUndoW(replaceTestStr)
                    If Not pIUtil.m_strcmpW(replaceTestStr.ToString(), replaceBeforeStr) Then
                        Throw New ArgumentException()
                    End If
                    pIUtil.m_replaceA(replaceTestStr)
                    If Not pIUtil.m_strcmpA(replaceTestStr.ToString(), replaceAfterStr) Then
                        Throw New ArgumentException()
                    End If
                    pIUtil.m_replaceUndoA(replaceTestStr)
                    If Not pIUtil.m_strcmpA(replaceTestStr.ToString(), replaceBeforeStr) Then
                        Throw New ArgumentException()
                    End If

                    pIUtil.m_regexReplaceW(regexTestNoDB, False)
                    If Not pIUtil.m_strcmpW(regexTestNoDB.ToString(), regexTestNoDBAfter) Then
                        Throw New ArgumentException()
                    End If
                    pIUtil.m_regexReplaceW(regexTestDB, True)
                    If Not pIUtil.m_strcmpW(regexTestDB.ToString(), regexTestDBAfter) Then
                        Throw New ArgumentException()
                    End If

                    'regex test
                    If Not pIUtil.m_regexMatchW(Unit_TestUpper, wildcard) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexMatchW(Unit_TestUpper, wildcardBeginUnit) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexMatchW(Unit_TestUpper, wildcardEndTest) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexMatchW(unit_test, wildcard) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_regexMatchW(unit_test, wildcardBeginUnit) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_regexMatchW(unit_test, wildcardEndTest) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_regexMatchW(unit_test, dotdotdot) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIUtil.m_regexMatchW(hi_, dotdotdot) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(hi_, dotdotdot) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(Unit_TestUpper, wildcard) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(Unit_TestUpper, wildcardBeginUnit) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(Unit_TestUpper, wildcardEndTest) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(unit_test, wildcard) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(unit_test, wildcardBeginUnit) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_regexiMatchW(unit_test, wildcardEndTest) Then
                        Throw New ArgumentException()
                    End If
                    If pIUtil.m_regexiMatchW(unit_test, dotdotdot) Then
                        Throw New ArgumentException()
                    End If
#End Region
#Region "formatVar___ functions"
                    Dim testVariant As Object() = New Object(10) {}
                    Dim outputString As New StringBuilder(&H512)
                    'TODO: Unable to force ansi string into object as it does not have support for it.
                    'testVariant[0] = new BStrWrapper("Aa"); //Nope, it's set to auto. It actually return unicode / System.Runtime.InteropServices.BStrWrapper
                    'testVariant[0] = Encoding.Default.GetString(Encoding.Default.GetBytes("Aa")); //Nope, still is unicode.
                    testVariant(1) = "aA"
                    testVariant(2) = CSng(1.0F)
                    testVariant(3) = CDbl(2.000002)
                    testVariant(4) = CBool(True)
                    testVariant(5) = CShort(25)
                    testVariant(6) = CUShort(50)
                    testVariant(7) = Int32.MaxValue 'MAXINT
                    testVariant(8) = UInt32.MaxValue 'MAXUINT
                    testVariant(9) = CLng(&H7FFFFFFF) 'MAXLONG
                    testVariant(10) = CULng(CUInt(Not CUInt(0))) 'MAXULONG
                    If Not pIUtil.m_formatVariantW(outputString, CUInt(outputString.Capacity), variantFormat, 11, testVariant) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIUtil.m_strcmpW(variantFormatExpected, outputString.ToString()) Then
                        Throw New ArgumentException()
                    End If
#End Region
                    MessageBox.Show("IUtil API has passed unit test.", "PASSED - IUtil", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("IUtil API has failed unit test.", "ERROR - IUtil", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "ICIniFile test section"
#If EXT_ICINIFILE Then
            Dim pICIniFile As Addon_API.ICIniFileClass = Addon_API.[Interface].getICIniFile(hash)
            Try
                If pICIniFile.isNotNull() Then
                    If pICIniFile.m_open_file(iniFileStr) Then
                        If Not pICIniFile.m_delete_file(iniFileStr) Then
                            Throw New ArgumentException()
                        End If
                        If pICIniFile.m_open_file(iniFileStr) Then
                            Throw New ArgumentException()
                        End If
                    End If
                    If Not pICIniFile.m_create_file(iniFileStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_open_file(iniFileStr) Then
                        Throw New ArgumentException()
                    End If
                    retCode = 0
recheckICIniFileDataExists:
                    If pICIniFile.m_section_exist(sectors.sect_name1) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name2) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name3) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name4) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name5) Then
                        Throw New ArgumentException()
                    End If

                    If pICIniFile.m_key_exist(sectors.sect_name1, str1_0) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name2, str1_1) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name3, str1_0) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name4, str1_2) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name5, str1_3) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_value_set(sectors.sect_name1, str1_0, firstUnitTestCStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_value_set(sectors.sect_name2, str1_1, firstUnitTestCStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_value_set(sectors.sect_name3, str1_0, firstUnitTestCStr) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_value_set(sectors.sect_name4, str1_2, firstUnitTestCStr) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_value_set(sectors.sect_name5, str1_3, firstUnitTestCStr) Then
                        Throw New ArgumentException()
                    End If
                    retCode += 1
                    Select Case retCode
                        Case 1
                            If Not pICIniFile.m_load() Then
                                Throw New ArgumentException()
                            End If
                            GoTo recheckICIniFileDataExists
                        Case 2
                            pICIniFile.m_clear()
                            If Not pICIniFile.m_save() Then
                                Throw New ArgumentException()
                            End If
                            If Not pICIniFile.m_load() Then
                                Throw New ArgumentException()
                            End If
                            GoTo recheckICIniFileDataExists
                        Case Else
                            Exit Select
                    End Select

                    If Not pICIniFile.m_save() Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_load() Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_section_exist(sectors.sect_name1) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_section_exist(sectors.sect_name2) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_section_exist(sectors.sect_name3) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name4) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name5) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_section_delete(sectors.sect_name3) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_section_exist(sectors.sect_name3) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_section_add(sectors.sect_name3) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_section_exist(sectors.sect_name3) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_key_exist(sectors.sect_name1, str1_0) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_key_exist(sectors.sect_name2, str1_1) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name3, str1_0) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name4, str1_2) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name5, str1_3) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_value_set(sectors.sect_name1, str1_0, firstUnitTestCStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICIniFile.m_key_exist(sectors.sect_name1, str1_0) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_save() Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_key_delete(sectors.sect_name1, str1_0) Then
                        Throw New ArgumentException()
                    End If
                    If pICIniFile.m_key_exist(sectors.sect_name1, str1_0) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_load() Then
                        Throw New ArgumentException()
                    End If

                    If Not pICIniFile.m_key_exist(sectors.sect_name1, str1_0) Then
                        Throw New ArgumentException()
                    End If

                    retCode = 1
                    Dim contentStr As String = Nothing
                    If Not pICIniFile.m_content(contentStr, retCode) Then
                        Throw New ArgumentException()
                    End If
                    If Not (contentStr IsNot Nothing AndAlso retCode <> 0) Then
                        Throw New ArgumentException()
                    End If

                    If iniFileDataStr.Length <> retCode Then
                        'Does not required -1 after Length
                        Throw New ArgumentException()
                    End If

                    'retCode++; //Is not required.

                    If Not compareString(contentStr, iniFileDataStr, retCode) Then
                        Throw New ArgumentException()
                    End If

                    ' Begin 0.5.3.4 Feature
                    Dim section_name As New StringBuilder(Addon_API.ICIniFileClass.INIFILESECTIONMAX)
                    Dim key_name As New StringBuilder(Addon_API.ICIniFileClass.INIFILEKEYMAX)
                    Dim value_name As New StringBuilder(Addon_API.ICIniFileClass.INIFILEVALUEMAX)
                    Dim ini_sec_count As UInteger = pICIniFile.m_section_count()
                    If ini_sec_count <> 3 Then
                        Throw New ArgumentException()
                    End If

                    Dim ini_key_count As UInteger
                    ' Section 0 test
                    If Not pICIniFile.m_section_index(0, section_name) Then
                        Throw New ArgumentException()
                    End If

                    If Not compareString(sectors.sect_name1, section_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If
                    ini_key_count = pICIniFile.m_key_count(section_name.ToString())
                    If ini_key_count <> 1 Then
                        Throw New ArgumentException()
                    End If

                    ' Section 0 key 0 test
                    If Not pICIniFile.m_key_index(section_name.ToString(), 0, key_name, value_name) Then
                        Throw New ArgumentException()
                    End If
                    If Not compareString(str1_0, key_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If
                    If Not compareString(firstUnitTestCStr, value_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If

                    ' Section 1 test
                    If Not pICIniFile.m_section_index(1, section_name) Then
                        Throw New ArgumentException()
                    End If

                    If Not compareString(sectors.sect_name2, section_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If
                    ini_key_count = pICIniFile.m_key_count(section_name.ToString())
                    If ini_key_count <> 1 Then
                        Throw New ArgumentException()
                    End If

                    ' Section 1 key 0 test
                    If Not pICIniFile.m_key_index(section_name.ToString(), 0, key_name, value_name) Then
                        Throw New ArgumentException()
                    End If
                    If Not compareString(str1_1, key_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If
                    If Not compareString(firstUnitTestCStr, value_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If

                    ' Section 2 test
                    If Not pICIniFile.m_section_index(2, section_name) Then
                        Throw New ArgumentException()
                    End If

                    If Not compareString(sectors.sect_name3, section_name.ToString(), UInteger.MaxValue) Then
                        Throw New ArgumentException()
                    End If
                    ini_key_count = pICIniFile.m_key_count(section_name.ToString())
                    If ini_key_count <> 0 Then
                        Throw New ArgumentException()
                    End If

                    ' End 0.5.3.4 Feature

                    If Not pICIniFile.m_delete_file(iniFileStr) Then
                        Throw New ArgumentException()
                    End If

                    pICIniFile.m_release()
                    MessageBox.Show("ICIniFile API has passed unit test.", "PASSED - ICIniFile", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                If pICIniFile.isNotNull() Then
                    pICIniFile.m_release()
                End If
                MessageBox.Show("ICIniFile API has failed unit test.", "ERROR - ICIniFile", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "ICommand test section"
#If EXT_ICOMMAND Then
            Dim pICommand As Addon_API.ICommand = Addon_API.[Interface].getICommand(hash)
            Try
                If pICommand.isNotNull() Then
                    'TODO: need to re-review this function internally.
                    If pICommand.m_reload_level(hash) Then
                        Throw New ArgumentException()
                    End If

                    'This is needed in order to preserve function pointer address
                    eao_testExecutePtr = AddressOf eao_testExecute
                    GC.KeepAlive(eao_testExecutePtr)
                    eao_testExecuteOverridePtr = AddressOf eao_testExecuteOverride
                    GC.KeepAlive(eao_testExecuteOverridePtr)
                    eao_testExecuteOverride2Ptr = AddressOf eao_testExecuteOverride2
                    GC.KeepAlive(eao_testExecuteOverride2Ptr)

                    If pICommand.m_delete(hash, eao_testExecutePtr, eaoTestExecuteStr) Then
                        Throw New ArgumentException()
                    End If
                    If pICommand.m_alias_delete(eaoTestExecuteStr, eaoTestExecuteAliasStr) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICommand.m_add(hash, eaoTestExecuteStr, eao_testExecutePtr, sectors.sect_name1, 1, 1,
                        False, HEXT.modeAll) Then
                        Throw New ArgumentException()
                    End If
                    If pICommand.m_add(hash, eaoTestExecuteStr, eao_testExecutePtr, sectors.sect_name1, 1, 1,
                        False, HEXT.modeAll) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICommand.m_delete(hash, eao_testExecutePtr, eaoTestExecuteStr) Then
                        Throw New ArgumentException()
                    End If
                    If pICommand.m_delete(hash, eao_testExecutePtr, eaoTestExecuteStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICommand.m_add(hash, eaoTestExecuteStr, eao_testExecutePtr, sectors.sect_name1, 1, 1,
                        True, HEXT.modeAll) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICommand.m_add(hash, eaoTestExecuteStr, eao_testExecuteOverridePtr, sectors.sect_name1, 1, 1,
                        True, HEXT.modeAll) Then
                        Throw New ArgumentException()
                    End If
                    If pICommand.m_add(hash, eaoTestExecuteStr, eao_testExecuteOverride2Ptr, sectors.sect_name1, 1, 1,
                        True, HEXT.modeAll) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICommand.m_alias_add(eaoTestExecuteStr, eaoTestExecuteAliasStr) Then
                        Throw New ArgumentException()
                    End If
                    If pICommand.m_alias_add(eaoTestExecuteStr, eaoTestExecuteAliasStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICommand.m_alias_delete(eaoTestExecuteStr, eaoTestExecuteAliasStr) Then
                        Throw New ArgumentException()
                    End If
                    If pICommand.m_alias_delete(eaoTestExecuteStr, eaoTestExecuteAliasStr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICommand.m_alias_add(eaoTestExecuteStr, eaoTestExecuteAliasStr) Then
                        Throw New ArgumentException()
                    End If

                    If Not pICommand.m_reload_level(hash) Then
                        Throw New ArgumentException()
                    End If
                    If Not pICommand.m_load_from_file(hash, eaoLoadFileStr, plI, Addon_API.MSG_PROTOCOL.MP_RCON) Then
                        Throw New ArgumentException()
                    End If

                    ' Proper remove command when done testing.
                    If Not pICommand.m_delete(hash, eao_testExecutePtr, eaoTestExecuteStr) Then
                        Throw New ArgumentException()
                    End If

                    MessageBox.Show("ICommand API has passed unit test.", "PASSED - ICommand", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("ICommand API has failed unit test.", "ERROR - ICommand", MessageBoxButtons.OK, MessageBoxIcon.[Error])

                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "IObject test section"
#If EXT_IOBJECT Then
            Dim pIObject As Addon_API.IObject = Addon_API.[Interface].getIObject(hash)
            Try
                If pIObject.isNotNull() Then
                    Dim gtag_list As New Addon_API.objTagGroupList()
                    If Not pIObject.m_get_lookup_group_tag_list(e_tag_group.TAG_WEAP, gtag_list) Then
                        Throw New ArgumentException()
                    End If
                    If gtag_list.count = 0 Then
                        Throw New ArgumentException()
                    End If
                    Dim tag_header As Addon_API.hTagHeader_managed = gtag_list.list(0)
                    If tag_header.isNull() Then
                        Throw New ArgumentException()
                    End If
                    If tag_header.hTagHeader_n.group_tag <> e_tag_group.TAG_WEAP Then
                        Throw New ArgumentException()
                    End If
                    Dim object_id As New s_ident(0)
                    Dim parent_id As New s_ident()
                    Dim move_object As New Addon_API.objManaged()
                    move_object.world.x = 1.0F
                    move_object.world.y = 1.0F
                    move_object.world.z = 1.0F
                    If Not pIObject.m_create(tag_header.hTagHeader_n.ident, parent_id, 1000, object_id, move_object.world) Then
                        Throw New ArgumentException()
                    End If
                    Dim created_object As s_object_managed = pIObject.m_get_address(object_id)
                    If created_object.getPtr().ptr = IntPtr.Zero Then
                        Throw New ArgumentException()
                    End If

                    tag_header = pIObject.m_lookup_tag(created_object.s_object_n.ModelTag)
                    If tag_header.isNull() Then
                        Throw New ArgumentException()
                    End If

                    If created_object.s_object_n.World.x <> 1.0F AndAlso created_object.s_object_n.World.y <> 1.0F AndAlso created_object.s_object_n.World.z <> 1.0F Then
                        Throw New ArgumentException()
                    End If
                    move_object.world.x = 2.0F
                    move_object.world.y = 2.0F
                    move_object.world.z = 2.0F
                    pIObject.m_move(object_id, move_object)
                    created_object.refresh()
                    If created_object.s_object_n.World.x <> 2.0F AndAlso created_object.s_object_n.World.y <> 2.0F AndAlso created_object.s_object_n.World.z <> 2.0F Then
                        Throw New ArgumentException()
                    End If
                    move_object.world.x = 5.0F
                    move_object.world.y = 5.0F
                    move_object.world.z = 5.0F
                    pIObject.m_move_and_reset(object_id, move_object.world)
                    created_object.refresh()
                    If created_object.s_object_n.World.x <> 5.0F AndAlso created_object.s_object_n.World.y <> 5.0F AndAlso created_object.s_object_n.World.z <> 5.0F Then
                        Throw New ArgumentException()
                    End If
                    pIObject.m_update(object_id)
                    If Not pIObject.m_destroy(object_id) Then
                        Throw New ArgumentException()
                    End If
                    MessageBox.Show("IObject API has passed unit test.", "PASSED - IObject", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("IObject API has failed unit test.", "ERROR - IObject", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "IPlayer test section"
#If EXT_IPLAYER AndAlso EXT_IADMIN = 0 Then
            Dim pIPlayer As Addon_API.IPlayer = Addon_API.[Interface].getIPlayer(hash)
            Try
                If pIPlayer.isNotNull() Then
                    Dim testStr As New StringBuilder(64)
                    Dim plList As New Addon_API.PlayerInfoList()

                    Dim totalPlayers As Short = pIPlayer.m_get_str_to_player_list("*", plList, Nothing)
                    If totalPlayers = 0 Then
                        Throw New ArgumentException()
                    End If
                    Dim plITest As New Addon_API.PlayerInfo(), plITest2 As New Addon_API.PlayerInfo()
                    If pIPlayer.m_get_m_index(2, plITest, True) Then
                        Throw New ArgumentException()
                    End If
                    If pIPlayer.m_get_m_index(1, plITest, True) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_m_index(0, plITest, True) Then
                        Throw New ArgumentException()
                    End If
                    If pIPlayer.m_get_id(200, plITest2) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_id(CUInt(plITest.plR.PlayerIndex), plITest2) Then
                        Throw New ArgumentException()
                    End If
                    If Not (plITest.cmS = plITest2.cmS AndAlso plITest.cplEx = plITest2.cplEx AndAlso plITest.cplS = plITest2.cplS AndAlso plITest.cplR = plITest2.cplR) Then
                        Throw New ArgumentException()
                    End If

                    Dim plBiped As s_biped_managed = pIObject.m_get_address(plITest.plS.CurrentBiped)
                    If plBiped.getPtr().ptr = IntPtr.Zero Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_ident(plBiped.s_object_n.PlayerOwner, plITest2) Then
                        Throw New ArgumentException()
                    End If

                    If pIPlayer.m_get_by_unique_id(600, plITest2) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_by_unique_id(plITest.mS.UniqueID, plITest2) Then
                        Throw New ArgumentException()
                    End If
                    If Not (plITest.cmS = plITest2.cmS AndAlso plITest.cplEx = plITest2.cplEx AndAlso plITest.cplS = plITest2.cplS AndAlso plITest.cplR = plITest2.cplR) Then
                        Throw New ArgumentException()
                    End If
                    retCode = pIPlayer.m_get_id_full_name(plITest.plR.PlayerName)
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_full_name_id(retCode, testStr) Then
                        Throw New ArgumentException()
                    End If
                    If testStr.ToString() <> plITest.plR.PlayerName Then
                        Throw New ArgumentException()
                    End If
                    testStr.Remove(0, testStr.Length)

                    retCode = pIPlayer.m_get_id_ip_address(plITest.plEx.IP_Addr)
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_ip_address_id(retCode, testStr) Then
                        Throw New ArgumentException()
                    End If
                    If testStr.ToString() <> plITest.plEx.IP_Addr Then
                        Throw New ArgumentException()
                    End If
                    testStr.Remove(0, testStr.Length)

                    retCode = pIPlayer.m_get_id_port(plITest.plEx.IP_Port)
                    If retCode = 0 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_port_id(retCode, testStr) Then
                        Throw New ArgumentException()
                    End If
                    If testStr.ToString() <> plITest.plEx.IP_Port Then
                        Throw New ArgumentException()
                    End If
                    testStr.Remove(0, testStr.Length)

                    If pIPlayer.m_update(plINull) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIPlayer.m_update(plITest) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIPlayer.m_send_custom_message(Addon_API.MSG_FORMAT.MF_BLANK, Addon_API.MSG_PROTOCOL.MP_CHAT, plITest, "Simple blank prefix message for {0:s}", 1, plITest.plR.PlayerName) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_send_custom_message(Addon_API.MSG_FORMAT.MF_SERVER, Addon_API.MSG_PROTOCOL.MP_CHAT, plITest, "Simple server prefix message for {0:s}", 1, plITest.plR.PlayerName) Then
                        Throw New ArgumentException()
                    End If

                    If Not pIPlayer.m_send_custom_message_broadcast(Addon_API.MSG_FORMAT.MF_BLANK, "Simple blank prefix message for {0:s}", 0) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_send_custom_message_broadcast(Addon_API.MSG_FORMAT.MF_SERVER, "Simple server prefix message for {0:s}", 0) Then
                        Throw New ArgumentException()
                    End If

                    'm_apply_camo test only required biped data to verify data is set to camoflauge.
                    plBiped.refresh()
                    If (plBiped.s_object_n.isVisible And &H10) <> 0 Then
                        Throw New ArgumentException()
                    End If
                    pIPlayer.m_apply_camo(plITest, 10)
                    plBiped.refresh()
                    If (plBiped.s_object_n.isVisible And &H10) = 0 Then
                        Throw New ArgumentException()
                    End If

                    Dim oldTeam As e_color_team_index = plITest.plR.Team
                    pIPlayer.m_change_team(plITest, CType(Convert.ToByte((oldTeam = e_color_team_index.TEAM_RED)), e_color_team_index), True)
                    If plITest.plR.Team = oldTeam Then
                        Throw New ArgumentException()
                    End If

                    Dim gmtm As New tm()
                    Dim time As System.DateTime = DateTime.UtcNow
                    gmtm.tm_isdst = Convert.ToInt32(time.IsDaylightSavingTime())
                    gmtm.tm_yday = time.DayOfYear
                    gmtm.tm_wday = CInt(time.DayOfWeek)
                    gmtm.tm_year = time.Year - 1900
                    gmtm.tm_mon = time.Month - 1
                    gmtm.tm_mday = time.Day
                    gmtm.tm_hour = time.Hour
                    gmtm.tm_min = time.Minute + 5
                    gmtm.tm_sec = time.Second
                    Dim plEx As Addon_API.PlayerExtended = plITest.plEx
                    If pIPlayer.m_ban_player(plEx, gmtm) = 0 Then
                        Throw New ArgumentException()
                    End If
                    Dim banID As UInteger, banID2 As UInteger 'Test CD hash key (un)ban verification
                    banID = pIPlayer.m_ban_CD_key_get_id(plITest.plEx.CDHashW)
                    If banID = 0 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_unban_id(banID) Then
                        Throw New ArgumentException()
                    End If
                    'TODO: Does not validate if CD hash is valid first before ban
                    If pIPlayer.m_ban_CD_key(plITest.plEx.CDHashW, gmtm) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    banID2 = pIPlayer.m_ban_CD_key_get_id(plITest.plEx.CDHashW)
                    If banID2 = 0 Then
                        Throw New ArgumentException()
                    End If
                    If banID <> banID2 Then
                        Throw New ArgumentException()
                    End If
                    'Test IP Address (un)ban verification
                    banID = pIPlayer.m_ban_ip_get_id(plITest.plEx.IP_Addr)
                    If banID = 0 Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_unban_id(banID) Then
                        Throw New ArgumentException()
                    End If
                    'TODO: Does not validate if IP Address is valid first before ban
                    If pIPlayer.m_ban_ip(plITest.plEx.IP_Addr, gmtm) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    banID2 = pIPlayer.m_ban_ip_get_id(plITest.plEx.IP_Addr)
                    If banID2 = 0 Then
                        Throw New ArgumentException()
                    End If
                    If banID <> banID2 Then
                        Throw New ArgumentException()
                    End If

                    Dim ipAddr As New in_addr()
                    Dim port As UShort = 0
                    Dim mS As s_machine_slot = plITest.mS
                    If Not pIPlayer.m_get_ip(mS, ipAddr) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_port(mS, port) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIPlayer.m_get_CD_hash(mS, testStr) Then
                        Throw New ArgumentException()
                    End If

                    'Uncomment this part if need to verify function return correctly with/out an admin player.
                    'if (!pIPlayer.m_is_admin((byte)mS.machineIndex))
                    'throw new ArgumentException();

                    MessageBox.Show("IPlayer API has passed unit test.", "PASSED - IPlayer", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("IPlayer API has failed unit test.", "ERROR - IPlayer", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "IAdmin test section"
#If EXT_IADMIN Then
#If EXT_IPLAYER = 0 Then
            Dim "EXT_IPLAYER Is required For testing EXT_IADMIN" As String
#End If
            Dim pIPlayer As Addon_API.IPlayer = Addon_API.[Interface].getIPlayer(hash)
            Dim pIAdmin As Addon_API.IAdmin = Addon_API.[Interface].getIAdmin(hash)
            Try
                If pIAdmin.isNotNull() AndAlso pIPlayer.isNotNull() Then
                    Dim plIMockup As New Addon_API.PlayerInfo()
                    If Not pIPlayer.m_get_m_index(0, plIMockup, True) Then
                        Throw New ArgumentException()
                    End If

                    If pIAdmin.m_is_username_exist(username) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_delete(username) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    Dim arg As New Addon_API.ArgContainer()
                    Dim func As Addon_API.CmdFunc = Nothing
                    Dim tmpLvl As Short = plIMockup.plEx.adminLvl
                    Dim tmpPlEx As Addon_API.PlayerExtended = plIMockup.plEx
                    tmpPlEx.adminLvl = 0
                    plIMockup.plEx = tmpPlEx
                    If pIAdmin.m_is_authorized(plIMockup, cmdEaoLoad, arg.vars, func) <> Addon_API.CMD_AUTH.DENIED Then
                        Throw New ArgumentException()
                    End If
                    tmpPlEx.adminLvl = 9999
                    plIMockup.plEx = tmpPlEx
                    If pIAdmin.m_is_authorized(plIMockup, cmdEaoLoad, arg.vars, func) <> Addon_API.CMD_AUTH.AUTHORIZED Then
                        Throw New ArgumentException()
                    End If
                    tmpPlEx.adminLvl = tmpLvl

                    If pIAdmin.m_add(tmpPlEx.CDHashW, tmpPlEx.IP_Addr, "0", username, username, 9999,
                        False, False) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_add(tmpPlEx.CDHashW, tmpPlEx.IP_Addr, "0", username, username, 9999,
                        False, False) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_is_username_exist(username) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_is_authorized(plIMockup, cmdEaoLoad, arg.vars, func) <> Addon_API.CMD_AUTH.AUTHORIZED Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_login(plIMockup, Addon_API.MSG_PROTOCOL.MP_CHAT, usernamebad, username) <> Addon_API.LOGIN_VALIDATION.FAIL Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_is_authorized(plIMockup, cmdEaoLoad, arg.vars, func) <> Addon_API.CMD_AUTH.DENIED Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_login(plIMockup, Addon_API.MSG_PROTOCOL.MP_CHAT, username, username) <> Addon_API.LOGIN_VALIDATION.OK Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_is_authorized(plIMockup, cmdEaoLoad, arg.vars, func) <> Addon_API.CMD_AUTH.AUTHORIZED Then
                        Throw New ArgumentException()
                    End If

                    If pIAdmin.m_delete(username) <> Addon_API.e_boolean.[TRUE] Then
                        Throw New ArgumentException()
                    End If
                    If pIAdmin.m_delete(username) <> Addon_API.e_boolean.[FALSE] Then
                        Throw New ArgumentException()
                    End If
                    MessageBox.Show("IAdmin API has passed unit test.", "PASSED - IAdmin", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("IAdmin API has failed unit test.", "ERROR - IAdmin", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "IHaloEngine test section"
#If EXT_IHALOENGINE Then
            Dim pIHaloEngine As Addon_API.IHaloEngine = Addon_API.[Interface].getIHaloEngine(hash)
            Try
                If pIHaloEngine.isNotNull() Then
                    Dim serverHeader As s_server_header_managed = pIHaloEngine.serverHeader
                    'TODO: Need find better test for serverHeader.
                    '                    if (serverHeader.data.totalPlayers !=1)
                    '                        throw new ArgumentException();
                    '                    

                    Dim playerReserved As s_player_reserved_slot_managed = pIHaloEngine.playerReserved
                    If Not (playerReserved.data.MachineIndex = 0 AndAlso playerReserved.data.PlayerIndex = 0) Then
                        Throw New ArgumentException()
                    End If

                    If pIHaloEngine.isDedi AndAlso pIHaloEngine.haloGameVersion = Addon_API.HALO_VERSION.CE Then
                        If pIHaloEngine.machineHeaderSize <> &HEC Then
                            Throw New ArgumentException()
                        End If
                    Else
                        If pIHaloEngine.machineHeaderSize <> &H60 Then
                            Throw New ArgumentException()
                        End If
                    End If
                    Dim mSIndex As s_machine_slot_managed = pIHaloEngine.machineHeader
                    If Not (mSIndex.data.machineIndex = 0 AndAlso mSIndex.data.isAvailable = 0 AndAlso mSIndex.data.data1 <> IntPtr.Zero AndAlso mSIndex.data.Unknown9 = &H7) Then
                        Throw New ArgumentException()
                    End If
                    Dim pl1 As New Addon_API.PlayerInfo()
                    pl1.cmS = mSIndex.getPtr().ptr
                    mSIndex += 1
                    If Not (mSIndex.data.machineIndex = -1 AndAlso mSIndex.data.data1 = IntPtr.Zero AndAlso mSIndex.data.Unknown9 = &H0) Then
                        Throw New ArgumentException()
                    End If
                    mSIndex -= 1
                    Dim mapHeader As s_map_header_managed = pIHaloEngine.mapCurrent
                    If mapHeader.data.head <> &H68656164 Then ''head'
                        Throw New ArgumentException()
                    End If
                    Select Case pIHaloEngine.haloGameVersion
                        Case Addon_API.HALO_VERSION.TRIAL
                            If mapHeader.data.haloVersion <> &H6 Then
                                Throw New ArgumentException()
                            End If
                            Exit Select
                        Case Addon_API.HALO_VERSION.CE
                            If mapHeader.data.haloVersion <> &H261 Then
                                Throw New ArgumentException()
                            End If
                            Exit Select
                        Case Addon_API.HALO_VERSION.PC
                            If mapHeader.data.haloVersion <> &H7 Then
                                Throw New ArgumentException()
                            End If
                            Exit Select
                        Case Addon_API.HALO_VERSION.UNKNOWN
                        Case Else
                            Throw New ArgumentException()
                    End Select
                    If pIHaloEngine.mapTimeLimitPermament.value <> UInt32.MaxValue Then
                        If pIHaloEngine.mapTimeLimitLive.value <> pIHaloEngine.mapTimeLimitPermament.value Then
                            Throw New ArgumentException()
                        End If
                    End If
                    Dim mapStatus As s_map_status_managed = pIHaloEngine.mapStatus
                    If mapStatus.data.upTime <> pIHaloEngine.mapUpTimeLive.value Then
                        Throw New ArgumentException()
                    End If
                    '
                    '                     * m_dispatch_rcon
                    '                     

                    Dim rcon As New rconDataManaged(rconTestStr)
                    pIHaloEngine.m_dispatch_rcon(rcon.data, pl1)
                    '
                    '                     * m_dispatch_player
                    '                     

                    Dim d As New chatDataManaged(playerChatTest, 0, chatType.TEAM)
                    ' Gotta pass a pointer to the chatData struct
                    Dim d_ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(chatData)))
                    Marshal.StructureToPtr(d.data, d_ptr, True)
                    ' Build the chat packet
                    Dim packetBuffer As Byte() = New Byte(4092 + (2 * playerChatTest.Length - 1)) {}
                    GC.KeepAlive(packetBuffer)
                    retCode = pIHaloEngine.m_build_packet(packetBuffer, 0, &HF, 0, d_ptr, 0,
                        1, 0)
                    pIHaloEngine.m_add_packet_to_player_queue(CUInt(mSIndex.data.machineIndex), packetBuffer, retCode, 1, 1, 0,
                        1, 3)
                    '
                    pIHaloEngine.m_dispatch_player(d.data, CUInt(Marshal.SizeOf(GetType(chatData))), pl1)
                    '
                    '                     * m_dispatch_global
                    '                     

                    d = New chatDataManaged(globalChatTest, 0, chatType.[GLOBAL])
                    ' Gotta pass a pointer to the chatData struct
                    Marshal.StructureToPtr(d.data, d_ptr, True)
                    'GC.KeepAlive(playerChatTest);
                    ' Build the chat packet
                    packetBuffer = New Byte(4092 + (2 * globalChatTest.Length - 1)) {}
                    GC.KeepAlive(packetBuffer)
                    retCode = pIHaloEngine.m_build_packet(packetBuffer, 0, &HF, 0, d_ptr, 0,
                        1, 0)
                    pIHaloEngine.m_add_packet_to_global_queue(packetBuffer, retCode, 1, 1, 0, 1, 3)
                    '
                    pIHaloEngine.m_dispatch_global(d.data, CUInt(Marshal.SizeOf(GetType(chatData))))
                    ' Since C# is a managed code, we need to free up allocated space.
                    Marshal.FreeHGlobal(d_ptr)
                    '
                    If pIHaloEngine.isDedi Then
                        If Not pIHaloEngine.m_map_next() Then
                            Throw New ArgumentException()
                        End If
                        If Not pIHaloEngine.m_set_idling() Then
                            Throw New ArgumentException()
                        End If
                    End If
                    If Not pIHaloEngine.m_send_reject_code(pIHaloEngine.machineHeader, Addon_API.REJECT_CODE.VIDEO_TEST) Then
                        Throw New ArgumentException()
                    End If
                    If Not pIHaloEngine.m_exec_command("sv_maplist") Then
                        Throw New ArgumentException()
                    End If
                    pIHaloEngine.m_set_server_password(password)
                    pIHaloEngine.m_get_server_password(passwordWGet)
                    If password <> passwordWGet.ToString() Then
                        Throw New ArgumentException()
                    End If
                    pIHaloEngine.m_set_rcon_password(password)
                    pIHaloEngine.m_get_rcon_password(passwordAGet)
                    If password <> passwordAGet.ToString() Then
                        Throw New ArgumentException()
                    End If

                    'Addon test section
                    'TODO: Both functions will not work in middle of load process.
                    '                    Addon_API.addon_info eaoInfo = new Addon_API.addon_info();
                    '                    if (!pIHaloEngine.m_ext_add_on_get_info_by_index(0, ref eaoInfo))
                    '                        throw new ArgumentException();
                    '                    eaoInfo.author.Remove(0, eaoInfo.author.Length);
                    '                    eaoInfo.config_folder.Remove(0, eaoInfo.config_folder.Length);
                    '                    if (!pIHaloEngine.m_ext_add_on_get_info_by_name(eaoInfo.name, ref eaoInfo))
                    '                        throw new ArgumentException();
                    '                    

                    'TODO: This function cannot be tested otherwise it will go in a loop plus is not fully implemented yet.
                    'if (!pIHaloEngine.m_ext_add_on_reload(EXTPluginInfo.name))
                    '    THROW(8);
                    MessageBox.Show("IHaloEngine API has passed unit test.", "PASSED - IHaloEngine", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("IHaloEngine API has failed unit test.", "ERROR - IHaloEngine", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
#Region "ITimer test section"
#If EXT_ITIMER Then
            pITimer = Addon_API.[Interface].getITimer(hash)
            Try
                If pITimer.isNotNull() Then

                    TimerID(0) = pITimer.m_add(hash, Nothing, 0) '1/30 second
                    If TimerID(0) = 0 Then
                        Throw New ArgumentException()
                    End If
                    TimerID(1) = pITimer.m_add(hash, Nothing, 60) '2 seconds
                    If TimerID(1) = 0 Then
                        Throw New ArgumentException()
                    End If
                    pITimer.m_delete(hash, TimerID(1))
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("ITimer API has failed unit test.", "ERROR - ITimer", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                Return Addon_API.EAO_RETURN.FAIL
            End Try
#End If
#End Region
            GC.Collect()
            Return Addon_API.EAO_RETURN.OVERRIDE
        End Function
        <DllExport("EXTOnEAOUnload", CallingConvention:=CallingConvention.Cdecl)>
        Public Shared Sub EXTOnEAOUnload()
        End Sub
#If EXT_HKTIMER Then
        <DllExport("EXTOnTimerExecute", CallingConvention:=CallingConvention.Cdecl)>
        Public Shared Function EXTOnTimerExecute(id As UInt32, count As UInt32) As <MarshalAs(UnmanagedType.I1)> Boolean
            Try
                If TimerID(0) = id Then
                    If TimerTickStart = 0 Then
                        TimerTickStart = CUInt(Environment.TickCount)
                        TimerID(0) = pITimer.m_add(hash, Nothing, 150) '5 seconds
                        If TimerID(0) = 0 Then
                            Throw New ArgumentException()
                        End If
                        TimerID(2) = pITimer.m_add(hash, Nothing, 30) '1 second
                        If TimerID(2) = 0 Then
                            Throw New ArgumentException()
                        End If
                    Else
                        TimerTickSys(0) = CUInt(Environment.TickCount)
                        Dim tmpTimerCheck As UInteger = TimerTickSys(0) - TimerTickStart
                        If Not (4500 < tmpTimerCheck AndAlso tmpTimerCheck < 5033) Then
                            Throw New ArgumentException()
                        End If
                        If TimerTickSys(1) <> 0 Then
                            Throw New ArgumentException()
                        End If
                        tmpTimerCheck = TimerTickSys(2) - TimerTickStart
                        If Not (500 < tmpTimerCheck AndAlso tmpTimerCheck < 1033) Then
                            Throw New ArgumentException()
                        End If
                        tmpTimerCheck = TimerTickSys(3) - TimerTickStart
                        If Not (2500 < tmpTimerCheck AndAlso tmpTimerCheck < 3033) Then
                            Throw New ArgumentException()
                        End If
                        MessageBox.Show("ITimer API has passed unit test.", "PASSED - ITimer", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                ElseIf TimerID(1) = id Then
                    TimerTickSys(1) = CUInt(Environment.TickCount)
                ElseIf TimerID(2) = id Then
                    TimerTickSys(2) = CUInt(Environment.TickCount)
                    TimerID(3) = pITimer.m_add(hash, Nothing, 60) '2 seconds
                    If TimerID(3) = 0 Then
                        Throw New ArgumentException()
                    End If
                ElseIf TimerID(3) = id Then
                    TimerTickSys(3) = CUInt(Environment.TickCount)
                Else
                    Throw New ArgumentException()
                End If
            Catch generatedExceptionName As ArgumentException
                MessageBox.Show("ITimer API has failed unit test.", "ERROR - ITimer", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End Try
            Return True
        End Function
        <DllExport("EXTOnTimerCancel", CallingConvention:=CallingConvention.Cdecl)>
        Public Shared Sub EXTOnTimerCancel(id As UInt32)
            If TimerID(0) = id Then
            ElseIf TimerID(1) = id Then
            ElseIf TimerID(2) = id Then
            ElseIf TimerID(3) = id Then
            Else
                MessageBox.Show("ITimer API has failed unit test.", "ERROR - ITimer", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End If
        End Sub
#End If
        Private Shared Function compareString(str1 As String, str2 As String, length As UInteger) As Addon_API.EAO_RETURN
            If length = UInteger.MaxValue Then
                length = 0
                If str1.Length <> str2.Length Then
                    Return False
                End If
                While str1.Length < length
                    If str1(CInt(length)) <> str2(CInt(length)) Then
                        Return False
                    End If
                End While
            Else
                Do
                    length -= 1
                    If str1(CInt(length)) <> str2(CInt(length)) Then
                        Return False
                    End If
                Loop While length <> 0
            End If
            Return True
        End Function
    End Class
End Namespace
