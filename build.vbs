Option Explicit

Const EXIT_SUCCESS = 0
Const EXIT_FAILURE = 1
Const ERR_ARGUMENTS = 1
Const ERR_BUILD = 2
Const ASCII_QUOTE = 34
Const SHELL_WINDOW_NORMAL = 1
Const PROCESS_RUNNING = 0
Const EXIT_POLL_MILLISECONDS = 50
Const CONSOLE_HOST_FILENAME = "cscript.exe"
Const SOURCE_DIRECTORY = "modules"
Const SOURCE_MANIFEST_FILENAME = "project.json"
Const BUILDER_RELATIVE_PATH = "tools\package_launcher.py"
Const DEFAULT_PACKAGE_RELATIVE_PATH = "build"
Const MAX_POSITIONAL_ARGUMENTS = 1
Const OUTPUT_DIRECTORY_ARGUMENT_INDEX = 0
Const BIN_DIRECTORY = "bin"
Const LAUNCHER_FILENAME = "launcher.vbs"
Const MACRO_FILENAME = "generator.swp"
Const BUILD_REPORT_FILENAME = "generator.build.json"
Const PYTHON_ENVIRONMENT_VARIABLE = "SW_DRAWING_PYTHON"
Const DEFAULT_PYTHON_COMMAND = "py -3"

Dim fso, shell, scriptDir

Sub Main()
    Dim result, errorNumber, errorText
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    If ForwardToConsole() Then
        WScript.Quit EXIT_SUCCESS
    End If
    On Error Resume Next
    result = BuildPackage()
    errorNumber = Err.Number
    errorText = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 Then
        WScript.Echo "BUILD ERROR " & errorNumber & ": " & errorText
        WScript.Quit EXIT_FAILURE
    End If
    WScript.Quit result
End Sub

Function BuildPackage()
    Dim named, key, packageRoot, builder, pythonPath, command, result, binDir
    Set named = WScript.Arguments.Named
    BuildPackage = EXIT_SUCCESS
    If named.Exists("help") Or named.Exists("?") Then
        WScript.Echo "Build a SolidWorks drawing-generator runtime package without starting CAD."
        WScript.Echo "cscript //nologo build.vbs [DIRECTORY]"
        WScript.Echo "Default package: build relative to this source checkout (ignored by Git)."
        WScript.Echo "An explicit relative directory is resolved from the current working directory."
        WScript.Echo "Output: bin\, bin.zip, and generator.build.json in the build directory. Python 3 is required."
        Exit Function
    End If
    For Each key In named
        Err.Raise vbObjectError + ERR_ARGUMENTS, , "Unknown build option: /" & key & ". Pass the output directory as a positional argument. Use /help."
    Next
    If WScript.Arguments.Unnamed.Count > MAX_POSITIONAL_ARGUMENTS Then
        Err.Raise vbObjectError + ERR_ARGUMENTS, , "Expected at most one output directory. Use /help."
    End If
    packageRoot = fso.GetAbsolutePathName(fso.BuildPath(scriptDir, DEFAULT_PACKAGE_RELATIVE_PATH))
    If WScript.Arguments.Unnamed.Count = MAX_POSITIONAL_ARGUMENTS Then
        If Len(Trim(WScript.Arguments.Unnamed(OUTPUT_DIRECTORY_ARGUMENT_INDEX))) = 0 Then
            Err.Raise vbObjectError + ERR_ARGUMENTS, , "A non-empty output directory is required."
        End If
        packageRoot = fso.GetAbsolutePathName(WScript.Arguments.Unnamed(OUTPUT_DIRECTORY_ARGUMENT_INDEX))
    End If
    builder = fso.BuildPath(scriptDir, BUILDER_RELATIVE_PATH)
    If Not fso.FileExists(fso.BuildPath(fso.BuildPath(scriptDir, SOURCE_DIRECTORY), SOURCE_MANIFEST_FILENAME)) Or Not fso.FileExists(builder) Then
        Err.Raise vbObjectError + ERR_BUILD, , "Build requires modules\project.json and tools\package_launcher.py in the source checkout."
    End If
    pythonPath = shell.Environment("PROCESS")(PYTHON_ENVIRONMENT_VARIABLE)
    command = DEFAULT_PYTHON_COMMAND
    If Len(pythonPath) > 0 Then
        command = Q(pythonPath)
    End If
    command = command & " -B " & Q(builder) & " " & Q(packageRoot)
    WScript.Echo "[build] Source: " & fso.BuildPath(scriptDir, SOURCE_DIRECTORY)
    WScript.Echo "[build] Output: " & packageRoot
    result = RunConsole(command)
    If result <> EXIT_SUCCESS Then
        WScript.Echo "[build] Failed; exit code=" & result
        BuildPackage = result
        Exit Function
    End If
    binDir = fso.BuildPath(packageRoot, BIN_DIRECTORY)
    For Each key In Array(LAUNCHER_FILENAME, MACRO_FILENAME)
        If Not fso.FileExists(fso.BuildPath(binDir, key)) Then
            Err.Raise vbObjectError + ERR_BUILD, , "Build did not produce " & fso.BuildPath(binDir, key)
        End If
    Next
    If Not fso.FileExists(fso.BuildPath(packageRoot, BUILD_REPORT_FILENAME)) Then
        Err.Raise vbObjectError + ERR_BUILD, , "Build did not produce " & fso.BuildPath(packageRoot, BUILD_REPORT_FILENAME)
    End If
    WScript.Echo "[build] Package ready: " & binDir
End Function

Function RunConsole(ByVal command)
    Dim child, errorNumber, errorText, details
    On Error Resume Next
    Set child = shell.Exec(command)
    errorNumber = Err.Number
    errorText = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 Then
        Err.Raise vbObjectError + ERR_BUILD, , "Python 3 is required to build the package. Install py -3 or set SW_DRAWING_PYTHON to python.exe. " & errorText
    End If
    Do While Not child.StdOut.AtEndOfStream
        WScript.Echo child.StdOut.ReadLine
    Loop
    details = child.StdErr.ReadAll
    If Len(Trim(details)) > 0 Then
        WScript.StdErr.Write details
    End If
    Do While child.Status = PROCESS_RUNNING
        WScript.Sleep EXIT_POLL_MILLISECONDS
    Loop
    RunConsole = child.ExitCode
End Function

Function ForwardToConsole()
    Dim command, arg
    ForwardToConsole = False
    If LCase(fso.GetFileName(WScript.FullName)) = CONSOLE_HOST_FILENAME Then
        Exit Function
    End If
    command = Q(fso.BuildPath(fso.GetParentFolderName(WScript.FullName), CONSOLE_HOST_FILENAME)) & " //nologo " & Q(WScript.ScriptFullName)
    For Each arg In WScript.Arguments
        command = command & " " & Q(arg)
    Next
    shell.Run command, SHELL_WINDOW_NORMAL, False
    ForwardToConsole = True
End Function

Function Q(ByVal value)
    Q = Chr(ASCII_QUOTE) & Replace(CStr(value), Chr(ASCII_QUOTE), Chr(ASCII_QUOTE) & Chr(ASCII_QUOTE)) & Chr(ASCII_QUOTE)
End Function

Main
