' We want to do several things with this file 
' 1. if the user enters a string, then set that as an out of office message 
' 2. if the user enters an email address, then redirect all mail to that user
' optionally send a copy to myself 

' https://technet.microsoft.com/en-us/library/ee692768.aspx  part 1
' https://technet.microsoft.com/en-us/library/ee692769.aspx  part 2

' initialise
Set fso = CreateObject("Scripting.FileSystemObject")
Set dctControlFile = CreateObject("Scripting.Dictionary")

' and then 
Set oShell = CreateObject( "WScript.Shell" )
strHhomeshare=oShell.ExpandEnvironmentStrings("%HOMESHARE%")

' we need these things global
strControlFile         = strHhomeshare & "\.forward"
strVacationFileDflt    = strHhomeshare & "\.vacation.msg"
Dim strVacationFile
Dim strForwardEmail
Dim strVacationMessage
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim arrControlFile(200) ' make sure it's big enough, memory is cheap
Dim intForwardLine
Dim intVacationSectionStart
Dim intVacationSubjectLine
Dim intVacationFileLine
Dim intVacationSectionEnd
Dim intControlFileEOF

Dim blOpSuccess

debugmode = false

' sendtoSelf.Checked = true

' and off we go...
Sub Window_onLoad
    window.resizeTo 700,500

    ' read the control file and message file, if they do not exist or are empty, read the template files instead
    If fso.FileExists(strControlFile) Then
        Set objControlFile = fso.GetFile(strControlFile)
        If objControlFile.Size > 0 Then 
            LoadControlFile 
        Else
            LoadDefaultControlFile
        End If
    Else
        LoadDefaultControlFile
    End If
    
    ' and now parse the array contents
    For i = 0 To intControlFileEOF
        strThisLine = arrControlFile(i)
        
        If InStr(1,strThisLine,"# Exim filter",vbTextCompare) Then blValidControlFile = true
        
        ' ==== Forward commands ============================================
        ' in this section we're interested in a line with a "deliver" command
        If InStr(1,strThisLine,"deliver",vbTextCompare) Then
            ' is it commented?
            strFwd1 = Left(strThisLine,1)
            If strFwd1 <> "#" Then radioRedirect.Checked = true
            ' split the line up on spaces, return all, vbTextCompare 
            arrForwardLine = split(strThisLine, " ", -1, vbTextCompare)
            ' the last one must be the address(es)
            strForwardEmail = arrForwardLine(Ubound(arrForwardLine))
            ' If Not (Instr(1, arrForwardLine(0), "#", vbTextCompare)) Then radioRedirect.Checked = true
            If Instr(1, strThisLine, "unseen", vbTextCompare) Then sendtoSelf.Checked = true
            ' we need this line for later
            intForwardLine = i
        End If
        
        ' ==== Vacation section ============================================
        If InStr(1,strThisLine,"if personal alias",vbTextCompare) Then 
            ' is it commented?
            strVac1 = Left(strThisLine,1)
            If strVac1 <> "#" Then radioVacation.Checked = true
            blVacationSection = true
            intVacationSectionStart = i
            strVacationLine = strThisLine
        End If
        
        ' ==== subject line
        If (blVacationSection And InStr(1,strThisLine,"subject",vbTextCompare)) Then 
            intStripLine = InStr(1,strThisLine,"subject",vbTextCompare) + Len("subject ")
            ' Replace(string,find,replacewith[,start[,count[,compare]]]) 
            ' return string starts from intStripLine, and then remove the quotes and any surrounding spaces
            intVacationSubjectLine = i
            strVacationSubject = Trim(Replace(strThisLine,"""","",intStripLine)) 
        End If
        
        ' ==== message file location
        If (blVacationSection And InStr(1,strThisLine,"file",vbTextCompare)) Then 
            intStripLine = InStr(1,strThisLine,"file",vbTextCompare) + Len("file ")
            strVacationFileLinuxVersion = Trim(Replace(strThisLine,"""","",intStripLine)) 
            intVacationFileLine = i
        End If
        
        ' ==== vacation section end
        If (blVacationSection And InStr(1,strThisLine,"endif",vbTextCompare)) Then 
            intVacationSectionEnd = i
            blVacationSection = false
        End If
       
    Next
    
    ' by now we have the linux version of where we think the vacation file should be, convert it to windows format
    strVacationFileWindowsVersion = Replace(strVacationFileLinuxVersion,"$home",strHhomeshare)
    strVacationFileWindowsVersion = Replace(strVacationFileWindowsVersion,"/","\")

    ' set strVacationFile to the referenced file or default
    If fso.FileExists(strVacationFileWindowsVersion) Then
        strVacationFile = strVacationFileWindowsVersion
    Else 
        strVacationFile = strVacationFileDflt
    End If
    
    ' default message
    LoadDefaultMessage
    ' now, if the file exists, and contains data, read it
    If fso.FileExists(strVacationFile) Then 
        Set objVacationFile = fso.GetFile(strVacationFile)
        If objVacationFile.Size > 0 Then
            Set objVacationFile = fso.OpenTextFile(strVacationFile, ForReading)
            strVacationMessage = objVacationFile.ReadAll
            objVacationFile.Close
        End If
    End If
    
    ' and preload the values into the text areas    
    ForwardAllTo.Value    = strForwardEmail
    VacationSubject.Value = strVacationSubject
    VacationMessage.Value = strVacationMessage
    
    ' =========== for debug ===========================================
    ' now process the file we read, rebuild the Control File for diags
    ' strForwardComands = strFwd1 & "intVacationSectionStart = " & intVacationSectionStart & " " & strVacationLine & vbCRLF & intVacationSectionEnd & "x: " & intStripLine & " """ & strVacationSubject & """" & vbCRLF
    ' strForwardComands = strForwardComands & 0 & " " & arrControlFile(0)
    ' For i = 1 To intControlFileEOF
        ' strForwardComands = strForwardComands & vbCRLF & i & " " & arrControlFile(i)
    ' Next
    ' strForwardComands = strVacationFileLinuxVersion & vbCRLF & strVacationFileWindowsVersion & vbCRLF & "strVacationFile = " & strVacationFile
    ' ControlFile.Value = strForwardComands
    
    ' one of the buttons should be checked
    If (Not (radioVacation.Checked Or radioRedirect.Checked)) Then radioClear.Checked = true
    

End Sub

Sub Submit
    For Each objButton in RadioOption
        If objButton.Checked Then
            button_selection = objButton.Value
        End If
    Next    
    
    Select Case button_selection
        Case "redirect" SetRedirect true
        Case "setooo"   SetVacation true
        Case "reset"    Reset
        Case Else Msgbox "Please select one of the radio buttons."
    End Select
    
    ' and write the control file
    If blOpSuccess Then 
        WriteControlFile
        Window.Close
    End If
End Sub

Sub Cancel
    ' no change
    Window.Close
End Sub

Sub CheckKey
    ' check every keypress, if it's Esc, then give up
    ' F1 if we want it: If window.event.keyCode = 112 Then self.close()
    If window.event.keyCode = 27 Then Window.Close
End Sub

Sub SetRedirect(blSet)
    ' sanity checking means that we ought to raise an error if blSet
    blOpSuccess = true
    If blSet Then
        strForwardEmail = ForwardAllTo.Value
        ' check for @ & . and the default address, start the compare from after 0, or a match will return 0 which looks like "not found"
        If InStr(1,strForwardEmail,"@",vbTextCompare) And InStr(1,strForwardEmail,".",vbTextCompare) And Len(strForwardEmail) > 10 _
            And InStr(1,strForwardEmail,"ecipient",vbTextCompare) = 0 And InStr(1,strForwardEmail,"company",vbTextCompare) = 0 Then 
            ' And (Not (InStr(1,strForwardEmail,"ecipient",vbTextCompare) Or InStr(1,strForwardEmail,"company.co.uk",vbTextCompare))) Then 
                                        arrControlFile(intForwardLine) = "deliver " & strForwardEmail
            If sendtoSelf.Checked Then  arrControlFile(intForwardLine) = "unseen "  & arrControlFile(intForwardLine)
        Else
            Msgbox "copy email address:" & vbCRLF & strForwardEmail & vbCRLF & "does not seem to be a valid email address"
            blOpSuccess = false
        End If
        ' and turn off any Vacation message
        SetVacation false
    Else 
        ' add a comment to the start of the line, unless it already has one
        If Left(arrControlFile(intForwardLine),1) <> "#" Then arrControlFile(intForwardLine) = "# " & arrControlFile(intForwardLine)
    End If

End Sub

Sub SetVacation(blSet)
    If blSet Then
        ' get the vacation file back into Linux format
        strVacationFileLinuxVersion = Replace(strVacationFile,strHhomeshare,"$home")
        strVacationFileLinuxVersion = Replace(strVacationFileLinuxVersion,"\","/")
    
        ' set subject and file
        arrControlFile(intVacationSubjectLine) = "   subject """ & VacationSubject.Value & """"
        arrControlFile(intVacationFileLine)    = "   file """ & strVacationFileLinuxVersion & """"
        ' uncomment any ooo settings in the array
        For i = intVacationSectionStart to intVacationSectionEnd
            If Left(arrControlFile(i),1) = "#" Then arrControlFile(i) = Mid(arrControlFile(i),3)
        Next

        ' and do this
        SetRedirect false

        ' write out the vacation file in case we changed it
        Set objVacationFile = fso.OpenTextFile(strVacationFile, ForWriting, true)
        objVacationFile.Write(VacationMessage.Value)
        Msgbox "Out of Office message set:" & vbCRLF & vbCRLF & arrControlFile(intVacationSubjectLine) & vbCRLF & VacationMessage.Value & vbCRLF & vbCRLF& "please remember to turn it off on your return"
    Else 
        ' add a comment to the start of the line, unless it already has one
        For i = intVacationSectionStart to intVacationSectionEnd
            If Left(arrControlFile(i),1) <> "#" Then arrControlFile(i) = "# " & arrControlFile(i)
        Next
    End If

End Sub

Sub Reset
    ' turn off Redirect and Vacation
    SetRedirect false
    SetVacation false
End Sub

Sub WriteControlFile
    ' if the file does not exist, we shall need to create it
    ' object.OpenTextFile (filename [, iomode[, createifnotexist[, format]]])
    Set objFile = fso.OpenTextFile(strControlFile, ForWriting, true)
    For i = Lbound(arrControlFile) To intControlFileEOF
        objFile.Writeline arrControlFile(i)
    Next
    objFile.Close
End Sub

Sub LoadControlFile
    Set objControlFile = fso.OpenTextFile(strControlFile, ForReading)
    i = 0
    ' and read the entire file into an array
    Do Until objControlFile.AtEndOfStream
        arrControlFile(i) = objControlFile.Readline
        intControlFileEOF = i
        i = i + 1
    Loop 
    objControlFile.Close
End Sub 

Sub LoadDefaultControlFile
    arrControlFile( 0) = "# Exim Filter"
    arrControlFile( 1) = "# the above line is required, it tells the system what syntax to expect in the file"
    arrControlFile( 2) = "# documentation at http://www.exim.org/exim-html-current/doc/html/spec_html/filter_ch-exim_filter_files.html"
    arrControlFile( 3) = "# folder names must have trailing /"
    arrControlFile( 4) = "# CONTAINS is a case sensitive match"
    arrControlFile( 5) = "# contains is a non case sensitive match"
    arrControlFile( 6) = ""
    arrControlFile( 7) = "# quit on failure"
    arrControlFile( 8) = "if error_message then finish endif"
    arrControlFile( 9) = ""
    arrControlFile(10) = "# this forwards the mail to someone@domain and optionally delivers it also your mailbox"
    arrControlFile(11) = "# unseen deliver recipient@company.co.uk"
    arrControlFile(12) = ""
    arrControlFile(13) = "# deal with Spam, these will be marked with"
    arrControlFile(14) = "# X-Spam-Level: YES or X-Spam-Status: Yes,  (be careful, this one contains baYES)"
    arrControlFile(15) = "if"
    arrControlFile(16) = "  $h_X-Spam-Level: contains ""YES"""
    arrControlFile(17) = "    or"
    arrControlFile(18) = "  $h_X-Spam-Status: CONTAINS ""Yes,"""
    arrControlFile(19) = "then"
    arrControlFile(20) = "  save $home/Maildir/.Junk/"
    arrControlFile(21) = "  finish"
    arrControlFile(22) = "endif"
    arrControlFile(23) = ""
    arrControlFile(24) = "# out of office"
    arrControlFile(25) = "if personal alias dave.evans@goodness.co.uk then"
    arrControlFile(26) = "    vacation to $reply_address"
    arrControlFile(27) = "    expand file $home/.vacation.msg"
    arrControlFile(28) = "    once $home/.vacation.db"
    arrControlFile(29) = "    log $home/.vacation.log"
    arrControlFile(30) = "    once_repeat 10d"
    arrControlFile(31) = "    from $local_part\@$domain"
    arrControlFile(32) = "    subject ""Auto: I am out of the office"""
    arrControlFile(33) = "endif"
    intControlFileEOF = 33
End Sub

Sub LoadDefaultMessage

    strVacationMessage   = "Subject: Auto: I away from the office" _
                & vbCRLF & "Precedence: bulk" _
                & vbCRLF & "This is an auto-response to your message: $SUBJECT" _
                & vbCRLF & "" _
                & vbCRLF & "I am currently out of the office." _
                & vbCRLF & "I will respond to your message on my return."
End Sub                






