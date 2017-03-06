' Bohack
' vWLC Putty Log Parser
' 2/28/17

Option Explicit
'Check for Arguments
If WScript.Arguments.Count = 0 Then
   Wscript.Echo "Usage: Script.vbs infile outfile"
   WScript.Quit
End If

Dim fso, inf, outf
Dim infile, outfile
Dim inline, outline
Dim rtime, macaddr, apname, apradioslot, clientstate, channel, currentrate, currentmode, ipaddress, ccxc, etwoe, signalstr, snr, accessvlan

Const fsoForWriting = 2

Dim mcslist, mcsarray, mcsnum
mcslist = "m0,6.5,m1,13,m2,19.5,m3,26,m4,39,m5,52,m6,58.5,m7,65,m8,13,m9,26,m10,39,m11,52,m12,78,m13,104,m14,117,m15,130,m16,19.5,m17,39,m18,58.5,m19,78,m20,117,m21,156,m22,175.5,m23,195,m0 ss1,6.5,m1 ss1,13,m2 ss1,19.5,m3 ss1,26,m4 ss1,39,m5 ss1,52,m6 ss1,58.5,m7 ss1,65,m8 ss1,78,m9 ss1,NA,m0 ss2,13,m1 ss2,26,m2 ss2,39,m3 ss2,52,m4 ss2,78,m5 ss2,104,m6 ss2,117,m7 ss2,130,m8 ss2,156,m9 ss2,78,m0 ss3,19.5,m1 ss3,39,m2 ss3,58.5,m3 ss3,78,m4 ss3,117,m5 ss3,156,m6 ss3,175.5,m7 ss3,195,m8 ss3,234,m9 ss3,260"
mcsarray = split(mcslist,",")

infile = WScript.Arguments(0)
outfile = WScript.Arguments(1)

Function GetType (ByVal mcsindex)
    For mcsnum = LBound(mcsarray) To UBound(mcsarray)
        If mcsindex = mcsarray(mcsnum) then GetType=mcsarray(mcsnum + 1): Exit Function
    Next
End Function

'Requires the SQLite3 ODBC Driver
Dim Conn, rs
Dim strSQL, oui
Set Conn = CreateObject("ADODB.Connection") 
Conn.Open "Driver={SQLite3 ODBC Driver};Database=oui.db;StepAPI=;Timeout=" 

Set fso = CreateObject("Scripting.FileSystemObject")
Set inf = fso.OpenTextFile(infile)
Set outf = fso.OpenTextFile(outfile, fsoForWriting, True)
outf.WriteLine "Time,Mac Address,OUI,AP Name,AP Radio Slot ID,Client State,Channel,Current Rate,Current Mode,IP Address,CCX Capability,Signal Strength,SNR,Access VLAN"
Do Until inf.AtEndOfStream
    inline = inf.ReadLine
    Select Case left(inline,49)
        Case "Time............................................."
            rtime = rtrim(mid(inline, 63, 8)
        Case "Client MAC Address..............................."
            macaddr = rtrim(mid(inline, 51, len(inline)-50))
            strSQL = "SELECT company from oui where mac='" & ucase(mid(replace(macaddr,":",""),1,6)) & "'"
            Set rs = Conn.Execute(strSQL)
            oui = replace(rs.getstring, chr(013),"")
            Set rs = Nothing
        Case "AP Name.........................................."
            apname = rtrim(mid(inline, 51, len(inline)-50))
        Case "AP radio slot Id................................."
            'AP Radio Slot is the radio 2.4ghz or 5ghz
            Select Case rtrim(mid(inline, 51, len(inline)-50))
                Case "0"
                    apradioslot = "2.4"
                Case "1"
                    apradioslot = "5"
            End Select
        Case "Client State....................................."
            clientstate = rtrim(mid(inline, 51, len(inline)-50))
        Case "Channel.........................................."
            channel = rtrim(mid(inline, 51, len(inline)-50))
        Case "Current Rate....................................."
            If mid(inline, 51, 1) = "m" Then
                currentrate = GetType(rtrim(mid(inline, 51, len(inline)-50)))
                If InStr(rtrim(mid(inline, 51, len(inline)-50)),"ss") Then
                    currentmode = "802.11ac"
                Else
                    currentmode = "802.11n"
                End If
            Else
                currentrate = rtrim(mid(inline, 51, len(inline)-50))
                Select Case apradioslot
                    Case "2.4"
                        currentmode = "802.11g"
                    Case "5"
                        currentmode = "802.11a"
                End Select
            End If
        Case "IP Address......................................."
            ipaddress = rtrim(mid(inline, 51, len(inline)-50))
        Case "Client CCX version..............................."
            ccxc = rtrim(mid(inline, 51, len(inline)-50))
        Case "Client E2E version..............................."
            etwoe = rtrim(mid(inline, 51, len(inline)-50))
        Case "      Radio Signal Strength Indicator............"
            signalstr = rtrim(mid(inline, 51, len(inline)-InStr(inline," dB")))
            If signalstr = "Unavailable" then signalstr = "NA"
        Case "      Signal to Noise Ratio......................"
            snr = rtrim(mid(inline, 51, len(inline)-InStr(inline," dB")))
            If snr = "Unavailable" then snr = "NA"
        Case "Access VLAN......................................"
            accessvlan = rtrim(mid(inline, 51, len(inline)-50))
        Case "Fastlane Client: ................................"
            outline = rtime & "," & macaddr & "," & oui &"," & apname & "," & apradioslot & "," & clientstate & "," & channel & "," & currentrate & "," & currentmode & "," & ipaddress & "," & ccxc & "," & etwoe & "," & signalstr & "," & snr & "," & accessvlan
            outf.WriteLine outline
    End select
Loop

outf.Close
inf.Close

Set Conn = Nothing
Set inf = Nothing
Set outf = Nothing
Set fso = Nothing