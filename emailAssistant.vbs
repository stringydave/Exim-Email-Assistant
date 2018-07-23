' We want to do several things with this file 
' 1. if the user enters an email address, then copy all mail to that user
' 2. if the user enters a vacation message, then set that as an out of office message 

' https://www.exim.org/exim-html-current/doc/html/spec_html/filter_ch-exim_filter_files.html
' https://technet.microsoft.com/en-us/library/ee692768.aspx  part 1
' https://technet.microsoft.com/en-us/library/ee692769.aspx  part 2

' todo: if the control file is not valid, then initialise it
' todo: cope with multiple vacation alias lines
' todo: trap failed to write to control file

' 09/04/18  dce  1.0 partly working, but lots more to do 
' 11/05/18  dce  1.2 vacation subject is no longer a thing, we hard code it
' 14/05/18  dce  1.3 add getUserDetails, get Username, Email Address from Thunderbird control files.
' 23/05/18  dce  1.4 correct typos in LoadDefaultControlFile
'                    use "unseen deliver"
' 23/07/18  dce  1.5 better code to locate where control files should go

' initialise
Set fso = CreateObject("Scripting.FileSystemObject")
Set dctControlFile = CreateObject("Scripting.Dictionary")

' and then we need to find out where the control files will be
Set oShell = CreateObject( "WScript.Shell" )
' the easiest one is %homeshare%, but if that's empty, then it will still contain "%" when we expand it
strHhomeshare=oShell.ExpandEnvironmentStrings("%HOMESHARE%")
' if that doesn't work try %homedrive% %homepath%
If InStr(1,strHhomeshare,"%",vbTextCompare) Then strHhomeshare=oShell.ExpandEnvironmentStrings("%HOMEDRIVE%") & oShell.ExpandEnvironmentStrings("%HOMEPATH%")
' if that doesn't work try %logonserver%\%username%, which may fail later if there's more than one %logonserver%
If InStr(1,strHhomeshare,"%",vbTextCompare) Then strHhomeshare=oShell.ExpandEnvironmentStrings("%LOGONSERVER%") & "\" & oShell.ExpandEnvironmentStrings("%USERNAME%")
' and if it's still empty, then give up.
If InStr(1,strHhomeshare,"%",vbTextCompare) Then 
    Msgbox "emailAssistant cannot work out where to save the control files" & vbCRLF & "files = " strHhomeshare & vbCRLF & "emailAssistant will now quit."
    Window.Close
End If

' we need these things global
strControlFile         = strHhomeshare & "\.forward"
strVacationFileDflt    = strHhomeshare & "\.vacation.msg"
Dim strMyName
Dim strMyEmailAddress
Dim strVacationFile
Dim strForwardEmail
Dim strVacationMessage
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim arrControlFile(200) ' make sure it's big enough, memory is cheap
Dim intForwardLine
Dim intUserNameLine
Dim intUserEmailLine
Dim intVacationSectionStart
Dim strVacationAlias 
Dim intVacationAliasLine 
Dim intVacationFileLine
Dim strVacationFrom
Dim intVacationFromLine
Dim intVacationSectionEnd
Dim intControlFileEOF

SetRedirectStatus = true 
SetVacationStatus = true

debugmode = false

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
    
    ' #### parse the array contents #####################################################
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
            ' we need this line for later
            intForwardLine = i
        End If
        
        ' ==== Vacation section ============================================
        If InStr(1,strThisLine,"if personal",vbTextCompare) Then 
            ' is it commented?
            strVac1 = Left(strThisLine,1)
            If strVac1 <> "#" Then radioVacation.Checked = true
            blVacationSection = true
            intVacationSectionStart = i
            strVacationLine = strThisLine
        End If
                
        ' ==== vacation alias
        If (blVacationSection And InStr(1,strThisLine,"alias",vbTextCompare)) Then 
            intStripLine = InStr(1,strThisLine,"alias",vbTextCompare) + Len("alias ")
            intVacationAliasLine = i
        End If
        
        ' ==== message file location
        If (blVacationSection And InStr(1,strThisLine,"file",vbTextCompare)) Then 
            intStripLine = InStr(1,strThisLine,"file",vbTextCompare) + Len("file ")
            strVacationFileLinuxVersion = Trim(Replace(strThisLine,"""","",intStripLine)) 
            intVacationFileLine = i
        End If
        
        ' ==== vacation from
        ' from "Firstname Secondname <recipient.name@company.co.uk>"

        If (blVacationSection And InStr(1,strThisLine,"from",vbTextCompare)) Then 
            intFirstQuote     = InStr(1,strThisLine,"""",vbTextCompare) + 1
            intFirstAngle     = InStr(1,strThisLine,"<",vbTextCompare)
            intSecondAngle    = InStr(1,strThisLine,">",vbTextCompare)
            ' only try to do the next bit if it's a properly formatted line
            If intFirstQuote <> 0 and intFirstAngle <> 0 and intSecondAngle <> 0 Then
                strMyName         = Trim(Mid(strThisLine,intFirstQuote,intFirstAngle - intFirstQuote))
                strMyEmailAddress = Trim(Mid(strThisLine,intFirstAngle + 1,intSecondAngle - intFirstAngle - 1))
            End If
            intVacationFromLine = i
        End If
        
        ' ==== vacation section end
        If (blVacationSection And InStr(1,strThisLine,"endif",vbTextCompare)) Then 
            intVacationSectionEnd = i
            blVacationSection = false
        End If
       
    Next
        
    ' if we have not yet retrieved the user details, go get them
    If InvalidEmail(strMyEmailAddress) or Len(strMyName & "") < 5 Then
        GetUserDetails
    End If
    
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
    MyName.Value          = strMyName
    MyEmailAddress.Value  = strMyEmailAddress
    VacationMessage.Value = strVacationMessage
    
    ' =========== for debug ===========================================
    ' now process the file we read, rebuild the Control File for diags
    ' strForwardComands = strFwd1 & "intVacationSectionStart = " & intVacationSectionStart & " " & strVacationLine & vbCRLF & intVacationSectionEnd & "x: " & intStripLine & vbCRLF
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
    If (SetRedirectStatus and SetVacationStatus) Then 
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
    SetRedirectStatus = true
    If blSet Then
        strForwardEmail = ForwardAllTo.Value
        If InvalidEmail(strForwardEmail) Then
            Msgbox "send a copy email address:" & vbCRLF & strForwardEmail & vbCRLF & "does not seem to be a valid email address"
            SetRedirectStatus = false
            Exit Sub
        Else
            arrControlFile(intForwardLine) = "unseen deliver " & strForwardEmail
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
    
        ' set file
        arrControlFile(intVacationFileLine)    = "  file """ & strVacationFileLinuxVersion & """"
        
        ' update the variables
        SetVacationStatus = true
        strMyName         = MyName.Value
        strMyEmailAddress = MyEmailAddress.Value
        If Len(strMyName) < 3 Then
            Msgbox "out of office ""my name""" & vbCRLF & strMyName & vbCRLF & "does not seem to be valid"
            SetVacationStatus = false
            Exit Sub
        End If
        
        If InvalidEmail(strMyEmailAddress) Then
            Msgbox "out of office ""my email address""" & vbCRLF & strMyEmailAddress & vbCRLF & "does not seem to be a valid email address"
            SetVacationStatus = false
            Exit Sub
        End If

        ' and we need to deal with the alias and from lines:
        ' alias recipient.name@company.co.uk"
        ' from "Firstname Secondname <recipient.name@company.co.uk>"
        
        ' todo: cope here with multiple alias lines

        arrControlFile(intVacationAliasLine) = "  alias " & strMyEmailAddress
        arrControlFile(intVacationFromLine)  = "  from """ & strMyName & " <" & strMyEmailAddress & ">"""
        
        ' uncomment any ooo settings in the array
        For i = intVacationSectionStart to intVacationSectionEnd
            If Left(arrControlFile(i),1) = "#" Then arrControlFile(i) = Mid(arrControlFile(i),3)
        Next

        ' and do this
        SetRedirect false

        ' write out the vacation file in case we changed it
        Set objVacationFile = fso.OpenTextFile(strVacationFile, ForWriting, true)
        objVacationFile.Write(VacationMessage.Value)
        Msgbox _
            "Out of Office message set:" & vbCRLF & vbCRLF & _
            "From: " & strMyName & " " & strMyEmailAddress & vbCRLF & _
            "Subject: Auto: Re: <message_subject>" & vbCRLF & _
            "Message: " & VacationMessage.Value & vbCRLF & vbCRLF& _
            "Please remember to turn this off on your return"
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
    arrControlFile(11) = "# unseen deliver recipient.name@company.co.uk"
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
    arrControlFile(25) = "# if personal"
    arrControlFile(26) = "#   alias "         ' leave blank to be filled in by GetUserDetails
    arrControlFile(27) = "#   then"
    arrControlFile(28) = "#   vacation to $reply_address"
    arrControlFile(29) = "#   expand file $home/.vacation.msg"
    arrControlFile(30) = "#   once $home/.vacation.db"
    arrControlFile(31) = "#   log $home/.vacation.log"
    arrControlFile(32) = "#   once_repeat 10d"
    arrControlFile(33) = "#   from """""       ' leave blank to be filled in by GetUserDetails
    arrControlFile(34) = "#   subject ""Auto: Re: $h_subject:"""
    arrControlFile(35) = "# endif"
    intControlFileEOF = 35
End Sub

Sub LoadDefaultMessage
    strVacationMessage   = "I am currently out of the office." _
                & vbCRLF & "I will respond to your message on my return."
End Sub                

Function InvalidEmail(strEmailAddress)
    ' the purpose of this function is to return check a string looks like a valid email address and true or false
    ' check for @ & . and the default address, start the compare from after 0, or a match will return 0 which looks like "not found"
    If Len(strEmailAddress) < 10 _
        Or InStr(1,strEmailAddress,"@",vbTextCompare) = 0 _
        Or InStr(1,strEmailAddress,".",vbTextCompare) = 0 _
        Or InStr(1,strEmailAddress,"ecipient",vbTextCompare) <> 0 _
        Or InStr(1,strEmailAddress,"company",vbTextCompare) <> 0 Then
        InvalidEmail = True
    Else
        InvalidEmail = False
    End If

End Function

Sub GetUserDetails
Msgbox "send a"
    ' called if the user details were not in the .forward file.  In our environment we can steal them from the 
    ' Thunderbird control file.  There's a file at:
    ' C:\Users\username\AppData\Roaming\Thunderbird\profiles.ini
    ' which tells us the location of the Thunderbird profile, in there is a prefs.js file, e.g. at
    ' C:\Users\username\AppData\Roaming\Thunderbird\Profiles\randomstring\prefs.js
    ' and we can read that to get the user full name and the email address.
    ' another way might be to do an LDAP query.

    ' and then 
    strUserProfile=oShell.ExpandEnvironmentStrings("%USERPROFILE%")

    ' we need these things
    strThunderbirdini      = strUserProfile & "\AppData\Roaming\Thunderbird\profiles.ini"
    Dim strThunderbirdPrefs
    Dim strIsRelative
    Dim strPath      
    Dim strDefault

    ' read the ini file, it contains text like:
    ' IsRelative=1
    ' Path=Profiles/randomstring
    ' Default=1

    If fso.FileExists(strThunderbirdini) Then
        Set objIniFile = fso.GetFile(strThunderbirdini)
        If objIniFile.Size > 0 Then 
            Set objIniFile = fso.OpenTextFile(strThunderbirdini, ForReading)
            ' and read the entire file
            Do Until (objIniFile.AtEndOfStream or strDefault = "1")
                strThisLine = objIniFile.Readline
                intEqualPos = InStr(1,strThisLine,"=",vbTextCompare) + 1
                ' Mid(string,start[,length]) 
                If InStr(1,strThisLine,"IsRelative",vbTextCompare) Then strIsRelative = Mid(strThisLine,intEqualPos) 'the string after the equals sign
                If InStr(1,strThisLine,"Path",vbTextCompare)       Then strPath       = Mid(strThisLine,intEqualPos) 'the string after the equals sign
                If InStr(1,strThisLine,"Default",vbTextCompare)    Then strDefault    = Mid(strThisLine,intEqualPos) 'the string after the equals sign
            Loop 
            objIniFile.Close
            if debugmode Then WScript.Echo "IsRelative " & strIsRelative
            if debugmode Then WScript.Echo "Path       " & strPath
            if debugmode Then WScript.Echo "Default    " & strDefault
        Else
            ' if we've failed, then crash out & rely on the user filling it in.
            Exit Sub
        End If
    End If

    ' now we've found out where the Profile is, we need to read the prefs.js file
    ' Replace(string,find,replacewith[,start[,count[,compare]]]) 
    If strIsRelative = "1" Then 
        strThunderbirdPrefs = strUserProfile & "\AppData\Roaming\Thunderbird\" & strPath & "\prefs.js"
    Else
        strThunderbirdPrefs = Replace(strPath & "\prefs.js","/","\",1,-1,vbTextCompare)
    End If

    ' turn the separators round the right way
        strThunderbirdPrefs = Replace(strThunderbirdPrefs,"/","\")
    if debugmode Then WScript.Echo "Path       " & strThunderbirdPrefs

    ' now we've found out where the Profile is, we need to read the prefs.js file, and find these lines
    ' user_pref("mail.identity.id1.fullName", "Firstname Secondname");
    ' user_pref("mail.identity.id1.useremail", "user.name@company.com");

    If fso.FileExists(strThunderbirdPrefs) Then
        Set objPrefsFile = fso.GetFile(strThunderbirdPrefs)
        If objPrefsFile.Size > 0 Then 
            Set objPrefsFile = fso.OpenTextFile(strThunderbirdPrefs, ForReading)
            ' and read the entire file
            Do Until (objPrefsFile.AtEndOfStream or strMyEmailAddress <> "")
                strThisLine = objPrefsFile.Readline
                intCommaPos = InStr(1,strThisLine,",",vbTextCompare) + 1
                ' we just want to read id1 even if there are many
                If InStr(1,strThisLine,"mail.identity.id1.fullName",vbTextCompare)   Then strMyName  = Mid(strThisLine,intCommaPos) 'the string after the comma
                If InStr(1,strThisLine,"mail.identity.id1.useremail",vbTextCompare)  Then strMyEmailAddress = Mid(strThisLine,intCommaPos) 'the string after the comma
            Loop 
            objPrefsFile.Close
            if debugmode Then WScript.Echo "strMyName  " & strMyName
            if debugmode Then WScript.Echo "strMyEmailAddress " & strMyEmailAddress
            
        Else
             WScript.Quit
        End If
    End If

    ' clean up the variables
    strMyName  =      Replace(strMyName,"""","")
    strMyName  = Trim(Replace(strMyName,");",""))
    strMyEmailAddress =      Replace(strMyEmailAddress,"""","")
    strMyEmailAddress = Trim(Replace(strMyEmailAddress,");",""))

End Sub




