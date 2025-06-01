' ============================================================================
' ULTIMATE VBScript DEMONSTRATION
' A comprehensive showcase of VBScript capabilities
' ============================================================================

Option Explicit

' Global variables
Dim objFSO, objShell, objNetwork
Dim strUserName, strComputerName, strCurrentDir

' Initialize objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")

' Get system information
strUserName = objNetwork.UserName
strComputerName = objNetwork.ComputerName
strCurrentDir = objShell.CurrentDirectory

' ============================================================================
' MAIN PROGRAM
' ============================================================================

Sub Main()
    Dim choice
    
    ShowWelcomeMessage
    
    Do
        choice = ShowMainMenu()
        
        Select Case choice
            Case "1"
                SystemInformation
            Case "2"
                FileOperations
            Case "3"
                MathCalculator
            Case "4"
                TextProcessor
            Case "5"
                DateTimeOperations
            Case "6"
                RegistryOperations
            Case "7"
                NetworkInformation
            Case "8"
                ProcessManager
            Case "9"
                CreateSampleFiles
            Case "10"
                ShowAnimatedMessage
            Case "0"
                ShowGoodbyeMessage
                Exit Do
            Case Else
                MsgBox "Invalid choice! Please select a valid option.", vbExclamation, "Error"
        End Select
    Loop
End Sub

' ============================================================================
' WELCOME AND MENU FUNCTIONS
' ============================================================================

Sub ShowWelcomeMessage()
    Dim welcomeMsg
    welcomeMsg = "========================================" & vbCrLf & _
                "    ULTIMATE VBScript DEMONSTRATION    " & vbCrLf & _
                "========================================" & vbCrLf & _
                "Welcome, " & strUserName & "!" & vbCrLf & _
                "Computer: " & strComputerName & vbCrLf & _
                "Current Directory: " & strCurrentDir & vbCrLf & _
                "========================================" & vbCrLf & _
                "This script demonstrates advanced VBScript capabilities!"
    
    MsgBox welcomeMsg, vbInformation, "Welcome"
End Sub

Function ShowMainMenu()
    Dim menuMsg
    menuMsg = "╔══════════════════════════════════════╗" & vbCrLf & _
              "║         MAIN MENU OPTIONS            ║" & vbCrLf & _
              "╠══════════════════════════════════════╣" & vbCrLf & _
              "║ 1  - System Information             ║" & vbCrLf & _
              "║ 2  - File Operations                ║" & vbCrLf & _
              "║ 3  - Math Calculator                ║" & vbCrLf & _
              "║ 4  - Text Processor                 ║" & vbCrLf & _
              "║ 5  - Date/Time Operations           ║" & vbCrLf & _
              "║ 6  - Registry Operations            ║" & vbCrLf & _
              "║ 7  - Network Information            ║" & vbCrLf & _
              "║ 8  - Process Manager                ║" & vbCrLf & _
              "║ 9  - Create Sample Files            ║" & vbCrLf & _
              "║ 10 - Animated Message               ║" & vbCrLf & _
              "║ 0  - Exit                           ║" & vbCrLf & _
              "╚══════════════════════════════════════╝" & vbCrLf & _
              "Enter your choice:"
    
    ShowMainMenu = InputBox(menuMsg, "Ultimate VBScript", "1")
End Function

' ============================================================================
' SYSTEM INFORMATION
' ============================================================================

Sub SystemInformation()
    Dim objWMI, colItems, objItem, sysInfo
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    
    sysInfo = "SYSTEM INFORMATION" & vbCrLf & String(50, "=") & vbCrLf
    
    For Each objItem in colItems
        sysInfo = sysInfo & "Computer Name: " & objItem.Name & vbCrLf
        sysInfo = sysInfo & "Manufacturer: " & objItem.Manufacturer & vbCrLf
        sysInfo = sysInfo & "Model: " & objItem.Model & vbCrLf
        sysInfo = sysInfo & "Total RAM: " & FormatNumber(objItem.TotalPhysicalMemory/1024/1024/1024, 2) & " GB" & vbCrLf
    Next
    
    ' Get OS Information
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each objItem in colItems
        sysInfo = sysInfo & "Operating System: " & objItem.Caption & vbCrLf
        sysInfo = sysInfo & "Version: " & objItem.Version & vbCrLf
        sysInfo = sysInfo & "Architecture: " & objItem.OSArchitecture & vbCrLf
    Next
    
    ' Get CPU Information
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Processor")
    For Each objItem in colItems
        sysInfo = sysInfo & "Processor: " & objItem.Name & vbCrLf
        sysInfo = sysInfo & "Cores: " & objItem.NumberOfCores & vbCrLf
    Next
    
    MsgBox sysInfo, vbInformation, "System Information"
End Sub

' ============================================================================
' FILE OPERATIONS
' ============================================================================

Sub FileOperations()
    Dim choice, fileName, fileContent
    
    choice = InputBox("File Operations:" & vbCrLf & _
                     "1 - Create Text File" & vbCrLf & _
                     "2 - Read Text File" & vbCrLf & _
                     "3 - List Directory Contents" & vbCrLf & _
                     "4 - Get File Properties" & vbCrLf & _
                     "Enter choice:", "File Operations", "1")
    
    Select Case choice
        Case "1"
            fileName = InputBox("Enter filename (with .txt extension):", "Create File", "sample.txt")
            fileContent = InputBox("Enter file content:", "File Content", "Hello from VBScript!")
            CreateTextFile fileName, fileContent
        Case "2"
            fileName = InputBox("Enter filename to read:", "Read File", "sample.txt")
            ReadTextFile fileName
        Case "3"
            ListDirectoryContents
        Case "4"
            fileName = InputBox("Enter filename for properties:", "File Properties", "sample.txt")
            GetFileProperties fileName
    End Select
End Sub

Sub CreateTextFile(fileName, content)
    Dim objFile
    Set objFile = objFSO.CreateTextFile(fileName, True)
    objFile.WriteLine content
    objFile.WriteLine "Created on: " & Now()
    objFile.WriteLine "Created by: Ultimate VBScript Demo"
    objFile.Close
    MsgBox "File '" & fileName & "' created successfully!", vbInformation, "File Created"
End Sub

Sub ReadTextFile(fileName)
    If objFSO.FileExists(fileName) Then
        Dim objFile, content
        Set objFile = objFSO.OpenTextFile(fileName, 1)
        content = objFile.ReadAll
        objFile.Close
        MsgBox "Content of '" & fileName & "':" & vbCrLf & String(30, "-") & vbCrLf & content, vbInformation, "File Content"
    Else
        MsgBox "File '" & fileName & "' not found!", vbExclamation, "File Not Found"
    End If
End Sub

Sub ListDirectoryContents()
    Dim objFolder, objFile, objSubFolder, content
    Set objFolder = objFSO.GetFolder(".")
    
    content = "DIRECTORY CONTENTS" & vbCrLf & String(40, "=") & vbCrLf & vbCrLf
    content = content & "FOLDERS:" & vbCrLf & String(20, "-") & vbCrLf
    
    For Each objSubFolder in objFolder.SubFolders
        content = content & "📁 " & objSubFolder.Name & vbCrLf
    Next
    
    content = content & vbCrLf & "FILES:" & vbCrLf & String(20, "-") & vbCrLf
    
    For Each objFile in objFolder.Files
        content = content & "📄 " & objFile.Name & " (" & FormatFileSize(objFile.Size) & ")" & vbCrLf
    Next
    
    MsgBox content, vbInformation, "Directory Contents"
End Sub

Function FormatFileSize(size)
    If size < 1024 Then
        FormatFileSize = size & " bytes"
    ElseIf size < 1048576 Then
        FormatFileSize = FormatNumber(size/1024, 2) & " KB"
    Else
        FormatFileSize = FormatNumber(size/1048576, 2) & " MB"
    End If
End Function

Sub GetFileProperties(fileName)
    If objFSO.FileExists(fileName) Then
        Dim objFile, props
        Set objFile = objFSO.GetFile(fileName)
        
        props = "FILE PROPERTIES" & vbCrLf & String(30, "=") & vbCrLf
        props = props & "Name: " & objFile.Name & vbCrLf
        props = props & "Size: " & FormatFileSize(objFile.Size) & vbCrLf
        props = props & "Created: " & objFile.DateCreated & vbCrLf
        props = props & "Modified: " & objFile.DateLastModified & vbCrLf
        props = props & "Accessed: " & objFile.DateLastAccessed & vbCrLf
        props = props & "Path: " & objFile.Path & vbCrLf
        
        MsgBox props, vbInformation, "File Properties"
    Else
        MsgBox "File '" & fileName & "' not found!", vbExclamation, "File Not Found"
    End If
End Sub

' ============================================================================
' MATH CALCULATOR
' ============================================================================

Sub MathCalculator()
    Dim num1, num2, operation, result
    
    num1 = CDbl(InputBox("Enter first number:", "Calculator", "10"))
    num2 = CDbl(InputBox("Enter second number:", "Calculator", "5"))
    operation = InputBox("Choose operation:" & vbCrLf & _
                        "+ (Addition)" & vbCrLf & _
                        "- (Subtraction)" & vbCrLf & _
                        "* (Multiplication)" & vbCrLf & _
                        "/ (Division)" & vbCrLf & _
                        "^ (Power)" & vbCrLf & _
                        "% (Modulus)", "Calculator", "+")
    
    Select Case operation
        Case "+"
            result = num1 + num2
        Case "-"
            result = num1 - num2
        Case "*"
            result = num1 * num2
        Case "/"
            If num2 <> 0 Then
                result = num1 / num2
            Else
                MsgBox "Error: Division by zero!", vbCritical, "Math Error"
                Exit Sub
            End If
        Case "^"
            result = num1 ^ num2
        Case "%"
            If num2 <> 0 Then
                result = num1 Mod num2
            Else
                MsgBox "Error: Modulus by zero!", vbCritical, "Math Error"
                Exit Sub
            End If
        Case Else
            MsgBox "Invalid operation!", vbExclamation, "Error"
            Exit Sub
    End Select
    
    MsgBox "Calculation Result:" & vbCrLf & _
           num1 & " " & operation & " " & num2 & " = " & result & vbCrLf & vbCrLf & _
           "Additional Info:" & vbCrLf & _
           "Square root of result: " & Sqr(Abs(result)) & vbCrLf & _
           "Absolute value: " & Abs(result), vbInformation, "Calculator Result"
End Sub

' ============================================================================
' TEXT PROCESSOR
' ============================================================================

Sub TextProcessor()
    Dim inputText, choice, result
    
    inputText = InputBox("Enter text to process:", "Text Processor", "Hello World! This is a VBScript demonstration.")
    
    choice = InputBox("Text Processing Options:" & vbCrLf & _
                     "1 - Convert to UPPERCASE" & vbCrLf & _
                     "2 - Convert to lowercase" & vbCrLf & _
                     "3 - Reverse Text" & vbCrLf & _
                     "4 - Count Characters/Words" & vbCrLf & _
                     "5 - Remove Spaces" & vbCrLf & _
                     "6 - Add Decorations", "Text Processor", "1")
    
    Select Case choice
        Case "1"
            result = UCase(inputText)
        Case "2"
            result = LCase(inputText)
        Case "3"
            result = ReverseString(inputText)
        Case "4"
            ShowTextStatistics inputText
            Exit Sub
        Case "5"
            result = Replace(inputText, " ", "")
        Case "6"
            result = "✨ " & inputText & " ✨"
    End Select
    
    MsgBox "Original: " & inputText & vbCrLf & vbCrLf & _
           "Processed: " & result, vbInformation, "Text Processing Result"
End Sub

Function ReverseString(str)
    Dim i, reversedStr
    reversedStr = ""
    For i = Len(str) To 1 Step -1
        reversedStr = reversedStr & Mid(str, i, 1)
    Next
    ReverseString = reversedStr
End Function

Sub ShowTextStatistics(text)
    Dim charCount, wordCount, sentences, vowels, consonants
    Dim i, char
    
    charCount = Len(text)
    wordCount = UBound(Split(text, " ")) + 1
    sentences = UBound(Split(text, ".")) + UBound(Split(text, "!")) + UBound(Split(text, "?"))
    
    vowels = 0
    consonants = 0
    
    For i = 1 To Len(text)
        char = LCase(Mid(text, i, 1))
        If InStr("aeiou", char) > 0 Then
            vowels = vowels + 1
        ElseIf char >= "a" And char <= "z" Then
            consonants = consonants + 1
        End If
    Next
    
    MsgBox "TEXT STATISTICS" & vbCrLf & String(30, "=") & vbCrLf & _
           "Characters: " & charCount & vbCrLf & _
           "Words: " & wordCount & vbCrLf & _
           "Sentences: " & sentences & vbCrLf & _
           "Vowels: " & vowels & vbCrLf & _
           "Consonants: " & consonants, vbInformation, "Text Statistics"
End Sub

' ============================================================================
' DATE/TIME OPERATIONS
' ============================================================================

Sub DateTimeOperations()
    Dim choice, result, inputDate, days
    
    choice = InputBox("Date/Time Operations:" & vbCrLf & _
                     "1 - Current Date/Time Info" & vbCrLf & _
                     "2 - Add Days to Date" & vbCrLf & _
                     "3 - Calculate Age" & vbCrLf & _
                     "4 - Days Until New Year" & vbCrLf & _
                     "5 - Format Date/Time", "Date/Time", "1")
    
    Select Case choice
        Case "1"
            ShowCurrentDateTime
        Case "2"
            inputDate = CDate(InputBox("Enter date (mm/dd/yyyy):", "Add Days", Date()))
            days = CInt(InputBox("Enter days to add:", "Add Days", "30"))
            result = DateAdd("d", days, inputDate)
            MsgBox "Original Date: " & inputDate & vbCrLf & _
                   "Days Added: " & days & vbCrLf & _
                   "New Date: " & result, vbInformation, "Date Calculation"
        Case "3"
            CalculateAge
        Case "4"
            ShowDaysUntilNewYear
        Case "5"
            ShowFormattedDateTime
    End Select
End Sub

Sub ShowCurrentDateTime()
    Dim info
    info = "CURRENT DATE/TIME INFORMATION" & vbCrLf & String(40, "=") & vbCrLf
    info = info & "Current Date: " & Date() & vbCrLf
    info = info & "Current Time: " & Time() & vbCrLf
    info = info & "Date/Time: " & Now() & vbCrLf
    info = info & "Day of Week: " & WeekdayName(Weekday(Date())) & vbCrLf
    info = info & "Day of Year: " & DatePart("y", Date()) & vbCrLf
    info = info & "Week of Year: " & DatePart("ww", Date()) & vbCrLf
    info = info & "Month: " & MonthName(Month(Date())) & vbCrLf
    info = info & "Year: " & Year(Date()) & vbCrLf
    
    MsgBox info, vbInformation, "Date/Time Information"
End Sub

Sub CalculateAge()
    Dim birthDate, age, days, months
    birthDate = CDate(InputBox("Enter your birth date (mm/dd/yyyy):", "Calculate Age", "01/01/1990"))
    
    age = DateDiff("yyyy", birthDate, Date())
    months = DateDiff("m", birthDate, Date())
    days = DateDiff("d", birthDate, Date())
    
    MsgBox "Age Calculation" & vbCrLf & String(20, "=") & vbCrLf & _
           "Birth Date: " & birthDate & vbCrLf & _
           "Current Date: " & Date() & vbCrLf & _
           "Age in Years: " & age & vbCrLf & _
           "Age in Months: " & months & vbCrLf & _
           "Age in Days: " & days, vbInformation, "Age Calculator"
End Sub

Sub ShowDaysUntilNewYear()
    Dim newYear, daysLeft
    newYear = CDate("01/01/" & (Year(Date()) + 1))
    daysLeft = DateDiff("d", Date(), newYear)
    
    MsgBox "Days until New Year " & (Year(Date()) + 1) & ": " & daysLeft & " days" & vbCrLf & _
           "Current Date: " & Date() & vbCrLf & _
           "New Year Date: " & newYear, vbInformation, "New Year Countdown"
End Sub

Sub ShowFormattedDateTime()
    Dim formats
    formats = "DATE/TIME FORMATS" & vbCrLf & String(30, "=") & vbCrLf
    formats = formats & "Long Date: " & FormatDateTime(Date(), 1) & vbCrLf
    formats = formats & "Short Date: " & FormatDateTime(Date(), 2) & vbCrLf
    formats = formats & "Long Time: " & FormatDateTime(Time(), 3) & vbCrLf
    formats = formats & "Short Time: " & FormatDateTime(Time(), 4) & vbCrLf
    formats = formats & "Custom: " & Format(Now(), "dddd, mmmm dd, yyyy 'at' hh:nn:ss AM/PM") & vbCrLf
    
    MsgBox formats, vbInformation, "Formatted Date/Time"
End Sub

' ============================================================================
' REGISTRY OPERATIONS (SAFE READ-ONLY)
' ============================================================================

Sub RegistryOperations()
    Dim choice
    
    choice = InputBox("Registry Operations (Read-Only):" & vbCrLf & _
                     "1 - Windows Version Info" & vbCrLf & _
                     "2 - System Environment" & vbCrLf & _
                     "3 - User Information", "Registry", "1")
    
    Select Case choice
        Case "1"
            ShowWindowsVersionInfo
        Case "2"
            ShowEnvironmentVariables
        Case "3"
            ShowUserRegistryInfo
    End Select
End Sub

Sub ShowWindowsVersionInfo()
    Dim version, productName, buildNumber, info
    On Error Resume Next
    
    version = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
    productName = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
    buildNumber = objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber")
    
    info = "WINDOWS VERSION INFORMATION" & vbCrLf & String(40, "=") & vbCrLf
    info = info & "Product Name: " & productName & vbCrLf
    info = info & "Version: " & version & vbCrLf
    info = info & "Build Number: " & buildNumber & vbCrLf
    
    MsgBox info, vbInformation, "Windows Version"
    On Error GoTo 0
End Sub

Sub ShowEnvironmentVariables()
    Dim envVars
    envVars = "ENVIRONMENT VARIABLES" & vbCrLf & String(30, "=") & vbCrLf
    envVars = envVars & "COMPUTERNAME: " & objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & vbCrLf
    envVars = envVars & "USERNAME: " & objShell.ExpandEnvironmentStrings("%USERNAME%") & vbCrLf
    envVars = envVars & "USERPROFILE: " & objShell.ExpandEnvironmentStrings("%USERPROFILE%") & vbCrLf
    envVars = envVars & "TEMP: " & objShell.ExpandEnvironmentStrings("%TEMP%") & vbCrLf
    envVars = envVars & "WINDIR: " & objShell.ExpandEnvironmentStrings("%WINDIR%") & vbCrLf
    envVars = envVars & "PROGRAMFILES: " & objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & vbCrLf
    
    MsgBox envVars, vbInformation, "Environment Variables"
End Sub

Sub ShowUserRegistryInfo()
    MsgBox "User registry information access requires elevated permissions." & vbCrLf & _
           "Current user: " & strUserName & vbCrLf & _
           "Computer: " & strComputerName, vbInformation, "User Info"
End Sub

' ============================================================================
' NETWORK INFORMATION
' ============================================================================

Sub NetworkInformation()
    Dim info
    info = "NETWORK INFORMATION" & vbCrLf & String(30, "=") & vbCrLf
    info = info & "User Name: " & objNetwork.UserName & vbCrLf
    info = info & "Computer Name: " & objNetwork.ComputerName & vbCrLf
    info = info & "User Domain: " & objNetwork.UserDomain & vbCrLf
    
    ' Get network adapters
    Dim objWMI, colItems, objItem
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    info = info & vbCrLf & "ACTIVE NETWORK ADAPTERS:" & vbCrLf & String(25, "-") & vbCrLf
    
    For Each objItem in colItems
        If Not IsNull(objItem.IPAddress) Then
            info = info & "Adapter: " & objItem.Description & vbCrLf
            info = info & "IP Address: " & objItem.IPAddress(0) & vbCrLf
            If Not IsNull(objItem.DefaultIPGateway) Then
                info = info & "Gateway: " & objItem.DefaultIPGateway(0) & vbCrLf
            End If
            info = info & vbCrLf
        End If
    Next
    
    MsgBox info, vbInformation, "Network Information"
End Sub

' ============================================================================
' PROCESS MANAGER
' ============================================================================

Sub ProcessManager()
    Dim choice
    choice = InputBox("Process Manager:" & vbCrLf & _
                     "1 - List Running Processes" & vbCrLf & _
                     "2 - Process Statistics" & vbCrLf & _
                     "3 - Start Calculator" & vbCrLf & _
                     "4 - Start Notepad", "Process Manager", "1")
    
    Select Case choice
        Case "1"
            ListRunningProcesses
        Case "2"
            ShowProcessStatistics
        Case "3"
            objShell.Run "calc.exe"
            MsgBox "Calculator started!", vbInformation, "Process Started"
        Case "4"
            objShell.Run "notepad.exe"
            MsgBox "Notepad started!", vbInformation, "Process Started"
    End Select
End Sub

Sub ListRunningProcesses()
    Dim objWMI, colItems, objItem, processList, count
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Process")
    
    processList = "RUNNING PROCESSES (Top 20)" & vbCrLf & String(40, "=") & vbCrLf
    count = 0
    
    For Each objItem in colItems
        If count < 20 Then
            processList = processList & objItem.Name & " (PID: " & objItem.ProcessId & ")" & vbCrLf
            count = count + 1
        End If
    Next
    
    MsgBox processList, vbInformation, "Running Processes"
End Sub

Sub ShowProcessStatistics()
    Dim objWMI, colItems, totalProcesses, totalMemory
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_Process")
    
    totalProcesses = 0
    totalMemory = 0
    
    For Each objItem in colItems
        totalProcesses = totalProcesses + 1
        If Not IsNull(objItem.WorkingSetSize) Then
            totalMemory = totalMemory + objItem.WorkingSetSize
        End If
    Next
    
    MsgBox "PROCESS STATISTICS" & vbCrLf & String(30, "=") & vbCrLf & _
           "Total Processes: " & totalProcesses & vbCrLf & _
           "Total Memory Usage: " & FormatNumber(totalMemory/1024/1024, 2) & " MB", _
           vbInformation, "Process Statistics"
End Sub

' ============================================================================
' SAMPLE FILE CREATION
' ============================================================================

Sub CreateSampleFiles()
    CreateHTMLFile
    CreateXMLFile
    CreateCSVFile
    CreateBatchFile
    
    MsgBox "Sample files created:" & vbCrLf & _
           "• sample.html - Web page" & vbCrLf & _
           "• data.xml - XML document" & vbCrLf & _
           "• report.csv - CSV spreadsheet" & vbCrLf & _
           "• demo.bat - Batch script", vbInformation, "Files Created"
End Sub

Sub CreateHTMLFile()
    Dim objFile, htmlContent
    Set objFile = objFSO.CreateTextFile("sample.html", True)
    
    htmlContent = "<!DOCTYPE html>" & vbCrLf & _
                  "<html>" & vbCrLf & _
                  "<head>" & vbCrLf & _
                  "    <title>VBScript Generated Page</title>" & vbCrLf & _
                  "    <style>body{font-family:Arial;background:linear-gradient(45deg,#FF6B6B,#4ECDC4);color:white;padding:20px;}</style>" & vbCrLf & _
                  "</head>" & vbCrLf & _
                  "<body>" & vbCrLf & _
                  "    <h1>🚀 Generated by Ultimate VBScript!</h1>" & vbCrLf & _
                  "    <p>This HTML file was created on " & Now() & "</p>" & vbCrLf & _
                  "    <p>User: " & strUserName & " | Computer: " & strComputerName & "</p>" & vbCrLf & _
                  "</body>" & vbCrLf & _
                  "</html>"
    
    objFile.Write htmlContent
    objFile.Close
End Sub

Sub CreateXMLFile()
    Dim objFile, xmlContent
    Set objFile = objFSO.CreateTextFile("data.xml", True)
    
    xmlContent = "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf & _
                 "<VBScriptDemo>" & vbCrLf & _
                 "    <GeneratedOn>" & Now() & "</GeneratedOn>" & vbCrLf & _
                 "    <User>" & strUserName & "</User>" & vbCrLf & _
                 "    <Computer>" & strComputerName & "</Computer>" & vbCrLf & _
                 "    <Features>" & vbCrLf & _
                 "        <Feature>File Operations</Feature>" & vbCrLf & _
                 "        <Feature>System Information</Feature>" & vbCrLf & _
                 "        <Feature>Math Calculator</Feature>" & vbCrLf & _
                 "        <Feature>Text Processing</Feature>" & vbCrLf & _
                 "        <Feature>Date/Time Operations</Feature>" & vbCrLf & _
                 "    </Features>" & vbCrLf & _
                 "</VBScriptDemo>"
    
    objFile.Write xmlContent
    objFile.Close
End Sub

Sub CreateCSVFile()
    Dim objFile, csvContent
    Set objFile = objFSO.CreateTextFile("report.csv", True)
    
    csvContent = "Name,Value,Type,Timestamp" & vbCrLf & _
                 """Computer Name""," & strComputerName & ",System," & Now() & vbCrLf & _
                 """User Name""," & strUserName & ",User," & Now() & vbCrLf & _
                 """Current Directory""," & strCurrentDir & ",Path," & Now() & vbCrLf & _
                 """Script Version"",Ultimate VBScript v1.0,Software," & Now() & vbCrLf & _
                 """Features"",10,Count," & Now()
    
    objFile.Write csvContent
    objFile.Close
End Sub

Sub CreateBatchFile()
    Dim objFile, batContent
    Set objFile = objFSO.CreateTextFile("demo.bat", True)
    
    batContent = "@echo off" & vbCrLf & _
                 "echo ==============================" & vbCrLf & _
                 "echo Ultimate VBScript Demo Batch" & vbCrLf & _
                 "echo ==============================" & vbCrLf & _
                 "echo Generated on: " & Now() & vbCrLf & _
                 "echo User: " & strUserName & vbCrLf & _
                 "echo Computer: " & strComputerName & vbCrLf & _
                 "echo." & vbCrLf & _
                 "echo Available files:" & vbCrLf & _
                 "dir *.html *.xml *.csv *.txt /b" & vbCrLf & _
                 "echo." & vbCrLf & _
                 "pause"
    
    objFile.Write batContent
    objFile.Close
End Sub

' ============================================================================
' ANIMATED MESSAGE
' ============================================================================

Sub ShowAnimatedMessage()
    Dim i, message, animatedMsg
    message = "Ultimate VBScript Demonstration!"
    
    For i = 1 To Len(message)
        animatedMsg = Left(message, i) & "_"
        CreateObject("WScript.Shell").PopUp animatedMsg, 1, "Animated Message", vbInformation
    Next
    
    ' Final message with decorations
    Dim finalMsg
    finalMsg = "🎉 " & String(40, "=") & " 🎉" & vbCrLf & _
               "    ULTIMATE VBSCRIPT DEMONSTRATION    " & vbCrLf & _
               "🎉 " & String(40, "=") & " 🎉" & vbCrLf & vbCrLf & _
               "✅ System Information Retrieval" & vbCrLf & _
               "✅ Advanced File Operations" & vbCrLf & _
               "✅ Mathematical Calculations" & vbCrLf & _
               "✅ Text Processing & Analysis" & vbCrLf & _
               "✅ Date/Time Manipulations" & vbCrLf & _
               "✅ Registry Operations (Safe)" & vbCrLf & _
               "✅ Network Information" & vbCrLf & _
               "✅ Process Management" & vbCrLf & _
               "✅ Dynamic File Generation" & vbCrLf & _
               "✅ Interactive User Interface" & vbCrLf & vbCrLf & _
               "🚀 VBScript - Powered by Windows Scripting! 🚀"
    
    MsgBox finalMsg, vbInformation, "Ultimate VBScript Features"
End Sub

' ============================================================================
' GOODBYE MESSAGE
' ============================================================================

Sub ShowGoodbyeMessage()
    Dim goodbyeMsg
    goodbyeMsg = "🌟 " & String(50, "=") & " 🌟" & vbCrLf & _
                 "        THANK YOU FOR USING ULTIMATE VBSCRIPT!" & vbCrLf & _
                 "🌟 " & String(50, "=") & " 🌟" & vbCrLf & vbCrLf & _
                 "📊 Session Summary:" & vbCrLf & _
                 "• User: " & strUserName & vbCrLf & _
                 "• Computer: " & strComputerName & vbCrLf & _
                 "• Session Time: " & Now() & vbCrLf & _
                 "• Current Directory: " & strCurrentDir & vbCrLf & vbCrLf & _
                 "🎯 What You Experienced:" & vbCrLf & _
                 "• Comprehensive VBScript capabilities" & vbCrLf & _
                 "• System integration and automation" & vbCrLf & _
                 "• File and data manipulation" & vbCrLf & _
                 "• Interactive user interfaces" & vbCrLf & _
                 "• Professional error handling" & vbCrLf & vbCrLf & _
                 "💡 VBScript Fun Facts:" & vbCrLf & _
                 "• Released by Microsoft in 1996" & vbCrLf & _
                 "• Based on Visual Basic language" & vbCrLf & _
                 "• Perfect for Windows automation" & vbCrLf & _
                 "• Still widely used in enterprises" & vbCrLf & vbCrLf & _
                 "🚀 Keep exploring and scripting!" & vbCrLf & _
                 "    Happy coding! 👨‍💻👩‍💻" & vbCrLf & vbCrLf & _
                 "✨ Goodbye from Ultimate VBScript! ✨"
    
    MsgBox goodbyeMsg, vbInformation, "Farewell!"
End Sub

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

Function GetCurrentDateTime()
    GetCurrentDateTime = Format(Now(), "yyyy-mm-dd hh:nn:ss")
End Function

Function CreateSeparator(length)
    CreateSeparator = String(length, "=")
End Function

Sub LogMessage(message)
    ' Simple logging function
    Dim logFile
    Set logFile = objFSO.OpenTextFile("vbscript_log.txt", 8, True)
    logFile.WriteLine GetCurrentDateTime() & " - " & message
    logFile.Close
End Sub

' ============================================================================
' ERROR HANDLING
' ============================================================================

Sub HandleError(errorSource)
    If Err.Number <> 0 Then
        MsgBox "Error in " & errorSource & ":" & vbCrLf & _
               "Number: " & Err.Number & vbCrLf & _
               "Description: " & Err.Description & vbCrLf & _
               "Source: " & Err.Source, vbCritical, "Error Occurred"
        Err.Clear
    End If
End Sub

' ============================================================================
' STARTUP
' ============================================================================

' Initialize error handling
On Error Resume Next

' Log script start
LogMessage "Ultimate VBScript Demo Started by " & strUserName

' Start main program
Main()

' Log script end
LogMessage "Ultimate VBScript Demo Ended"

' Cleanup objects
Set objFSO = Nothing
Set objShell = Nothing
Set objNetwork = Nothing

' Final cleanup message
WScript.Echo "🎉 Ultimate VBScript Demo completed successfully! 🎉" & vbCrLf & _
             "Check the generated files in your current directory!" & vbCrLf & _
             "Log file: vbscript_log.txt"

' End of script
