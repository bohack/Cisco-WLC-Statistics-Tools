' Bohack
' vWLC Putty Log Parser
' 3/15/17

Option Explicit
'Check for Arguments
If WScript.Arguments.Count = 0 Then
   Wscript.Echo "Usage: Script.vbs infile outfile"
   WScript.Quit
End If

Dim fso, inf, outf
Dim infile, outfile
Dim inline, outline
Dim rtime, macaddr, apname, apradioslot, clientstate, channel, currentrate, currentmode, preamble, ipaddress, ccxc, etwoe, signalstr, snr, accessvlan

Const fsoForWriting = 2

Dim mcslist, mcsarray, mcsnum, mcsmultipler, mcsvalue
mcslist = "m0,6.5,7.2,m1,13,14.4,m2,19.5,21.7,m3,26,28.9,m4,39,43.3,m5,52,57.8,m6,58.5,65,m7,65,72.2,m8,13,14.4,m9,26,28.9,m10,39,43.3,m11,52,57.8,m12,78,86.7,m13,104,115.6,m14,117,130,m15,130,144.4,m16,19.5,21.7,m17,39,43.3,m18,58.5,65,m19,78,86.7,m20,117,130,m21,156,173.3,m22,175.5,195,m23,195,216.7,m0 ss1,6.5,7.2,m1 ss1,13,14.4,m2 ss1,19.5,21.7,m3 ss1,26,28.9,m4 ss1,39,43.3,m5 ss1,52,57.8,m6 ss1,58.5,65,m7 ss1,65,72.2,m8 ss1,78,86.7,m9 ss1,NA,NA,m0 ss2,13,14.4,m1 ss2,26,28.9,m2 ss2,39,43.3,m3 ss2,52,57.8,m4 ss2,78,86.7,m5 ss2,104,115.6,m6 ss2,117,130,m7 ss2,130,144.4,m8 ss2,156,173.3,m9 ss2,78,NA,m0 ss3,19.5,21.7,m1 ss3,39,43.3,m2 ss3,58.5,65,m3 ss3,78,86.7,m4 ss3,117,130,m5 ss3,156,173.3,m6 ss3,175.5,195,m7 ss3,195,216.7,m8 ss3,234,260,m9 ss3,260,288.9"
mcsarray = split(mcslist,",")

infile = WScript.Arguments(0)
outfile = WScript.Arguments(1)

Function GetType (ByVal mcsindex, ByVal mcsmultipler)
    For mcsnum = LBound(mcsarray) To UBound(mcsarray)
        If mcsindex = mcsarray(mcsnum) then GetType=mcsarray(mcsnum + mcsmultipler): Exit Function
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
outf.WriteLine "Time,Mac Address,OUI,AP Name,AP Radio Slot ID,Client State,Channel,Current Rate,Current Mode,Preamble,IP Address,CCX Capability,E2E Capability,Signal Strength,SNR,Access VLAN"
Do Until inf.AtEndOfStream
    inline = inf.ReadLine
    Select Case left(inline,49)
        Case "Time............................................."
            rtime = rtrim(mid(inline, 62, 8))
        Case "Client MAC Address..............................."
            macaddr = rtrim(mid(inline, 51, len(inline)-50))
            strSQL = "SELECT company from oui where mac='" & ucase(mid(replace(macaddr,":",""),1,6)) & "'"
            Set rs = Conn.Execute(strSQL)
            IF rs.EOF = true Then
               oui = "Not Found"
            Else
               oui = replace(rs.getstring, chr(013),"")
            End If 
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
                mcsvalue = rtrim(mid(inline, 51, len(inline)-50))
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
        Case "      Short Preamble............................."
            Preamble = rtrim(mid(inline, 51, len(inline)-50))
            If currentmode = "802.11ac" or currentmode = "802.11n" Then
               If Preamble = "Implemented" Then
                  currentrate = GetType(mcsvalue,2)
               Else
                  currentrate = GetType(mcsvalue,1)
               End If
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
            outline = rtime & "," & macaddr & "," & oui &"," & apname & "," & apradioslot & "," & clientstate & "," & channel & "," & currentrate & "," & currentmode & "," & preamble & "," & ipaddress & "," & ccxc & "," & etwoe & "," & signalstr & "," & snr & "," & accessvlan
            outf.WriteLine outline
    End select
Loop

outf.Close
inf.Close

Set Conn = Nothing
Set inf = Nothing
Set outf = Nothing
Set fso = Nothing
