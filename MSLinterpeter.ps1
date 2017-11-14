#THIS IS THE INTERPRETER FOR MY VERY OWN SCRIPT LANGUAGE
# make .msl the extension (ie My Script Language - M.S.L)
#and also an excuse to learn powershell and to try to automate various gamming functions
#========================

param ([string]$targetScript)

$targetScript

[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
$DebugViewWindow_TypeDef = @'
[DllImport("user32.dll")]
public static extern IntPtr FindWindow(string ClassName, string Title);
[DllImport("user32.dll")]
public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")]
public static extern bool SetCursorPos(int X, int Y);
[DllImport("user32.dll")]
public static extern bool GetCursorPos(out System.Drawing.Point pt);
 
[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
 
private const int MOUSEEVENTF_LEFTDOWN = 0x02;
private const int MOUSEEVENTF_LEFTUP = 0x04;
private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
private const int MOUSEEVENTF_RIGHTUP = 0x10;
 
public static void LeftClick(){
    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
}
 
public static void LeftMouseDown2(){
    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
}
 
public static void LeftMouseUp2(){
    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
}
 
public static void RightClick(){
    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
}
'@
Add-Type -MemberDefinition $DebugViewWindow_TypeDef -Namespace AutoClicker2 -Name Temp2 -ReferencedAssemblies System.Drawing

#Add-Type -Namespace MyKeys -Name SendKeys -AssemblyName System.Windows.Forms
 
$mousePt = New-Object System.Drawing.Point
 ([AutoClicker2.Temp2]::GetCursorPos([ref]$mousePt)) | Out-Null

function normaliseVerb ($theLine) {
 #parse just for the verb
 $verb = $theLine.Split(" ")[0].Trim()
 if ($verb -like "//*") {$verb = "COMMENT"}
 elseif ($verb -like "#*") {$verb = "COMMENT"}
 elseif ($verb -eq "") {$verb = "COMMENT"}  # we ignore empty lines (just like comments) so call it a comment
 else {}
 $verb

}
#========================

function handleRUN ($theLine) {
 $LRunFile = $theLine.Remove(0, 4).Trim() #remove the verb
  start-process -FilePath $LRunFile 

}

#========================
function handleRunSubScript ($theLine) {
 #run a sub script
 #we need an array holding [curln#, maxln#, theScript] for each script depth
 #so look up arrays
 $LRunFile = $theLine.Remove(0, 12).Trim() #remove the verb

}

#========================
function handleRunWithArgs ($theLine) {
 $LRunFile = $theLine.Remove(0, 11).Trim() #remove the verb
 $theLine = $LRunFile
 $LArray = $LRunFile.indexof(" -a")
 $LRunFile = $theLine.substring(0, $LArray).Trim()
 $LArgs = $theLine.substring(3+$LArray,-3+-$LArray+$theline.length).Trim()

 start-process -FilePath $LRunFile -ArgumentList $LArgs

}
#========================

function handleRunWithInputFile ($theLine) {
 $LRunFile = $theLine.Remove(0, 16).Trim() #remove the verb
 $theLine = $LRunFile
 $LArray = $LRunFile.indexof(" -inputfile")
 $LRunFile = $theLine.substring(0, $LArray).Trim()
 $LArgs = $theLine.substring(11+$LArray,-11+-$LArray+$theline.length).Trim()
 $LRunFile
 $LArgs

 start-process -FilePath $LRunFile -RedirectStandardInput $LArgs

}
#========================

function handleRunAsAdmin ($theLine) {

 $LRunFile = $theLine.Remove(0, 10).Trim() #remove the verb
 start-process -Verb runas -FilePath $LRunFile

}
#========================

function LLSetMousePt ($X, $Y) {
 #NOT WORKING - KEEPS CLAIMING IT CANT CONVERT TO PARMAMS TO WAT setCursorPos needs??
 #move the mouse to the given pt (dont click just move)
 [AutoClicker2.Temp2]::SetCursorPos($X, $Y) | Out-Null

}
#========================

function handleSetMousePt ($thePt) {
 #move the mouse to the given pt (dont click just move)
 $LRunFile = $thePt.Remove(0, 8).Trim() #remove the verb
 $LArray = $LRunFile.Split(",").Trim()
 $LX = $LArray[0]
 $LY = $LArray[1]
 $mousePt.X = $LX
 $mousePt.Y = $LY

 [AutoClicker2.Temp2]::SetCursorPos($mousePt.X, $mousePt.Y) | Out-Null

}
#========================

function handleLeftClick ($theLine) {
 #click the mouse at the last specified pt (specified by SetMouse)
  [AutoClicker2.Temp2]::LeftClick() | Out-Null

}
#========================

function handleWaitSeconds ($theLine) {
 #sleep for the given number of seconds
 $LRunFile = $theLine.Remove(0, 4).Trim() #remove the verb
 Start-Sleep -Seconds $LRunFile

}

#========================
function handleWriteLine ($theLine) {
 #writeLine to host
 $LRunFile = $theLine.Remove(0, 7) #remove the verb
 Write-Host $LRunFile

}

#========================
function handleSendKeys ($theLine) {
 #send keys to the active window (or the void if no active window)
 $LRunFile = $theLine.Remove(0, 9) #remove the verb + 1 space
 [System.Windows.Forms.SendKeys]::SendWait($LRunFile);

}
#========================

function handleCloseWindowName ($theLine) {
 #kill the process with given name
 $LRunFile = $theLine.Remove(0, 15).Trim() #remove the verb
 #the following handles closing a process that does not exist silently (ie without errors)
 get-process * | Where-Object -Property Name -eq $LRunFile | Stop-Process -ErrorAction Stop -WarningAction Stop -Force  #NB: im trying to avoid using -Force (that would be dangerous!!!)

}
#========================

#we need to pass the script as an arg  Q: how?
#$inputFileName = "C:\Users\sauron\Documents\Programming\Powershell\MyScriptLanguage\testMSL.msl"
$inputFileName = $targetScript
Set-Location C:
$logFileName = ".\logfile.txt"
remove-item $logFileName | out-null
New-Item $logFileName -ItemType file | Out-Null



$log = [System.IO.StreamWriter] $logFileName 

 
$theScript = Get-Content $inputfilename
#main loop - ie read and interpret the script
$scriptDepth = 1    #init to the first index in an array (is that 0 or 1?)
$lineNumber = 0
$maxLineNumber = $theScript.Length

while ($lineNumber -lt $maxLineNumber) {
 #parse the line and get the verb/command in normalised form
 $theLine = $theScript[$lineNumber]
 $verb = normaliseVerb($theLine)

 if ($verb -imatch "COMMENT") {
  #no action required with comments
 } 
 elseif ($verb -imatch "EXIT") {
  #exit/terminate the script
  $log.WriteLine("EXIT")
  $lineNumber = $maxLineNumber    #crude but effective
 }
 elseif ($verb -imatch "RUNASADMIN") {
  #run an external program
  $log.WriteLine("RUN AS ADMIN")
  handleRunAsAdmin($theLine)
 }
 elseif ($verb -imatch "RUNWITHARGS") {
  #run an external program
  $log.WriteLine("RUN WITH ARGS")
  handleRunWithArgs($theLine)
 }
 elseif ($verb -imatch "RUNWITHINPUTFILE") {
  #run an external program
  $log.WriteLine("RUN WITH INPUTFILE")
  handleRunWithInputFile($theLine)
 }
 elseif ($verb -imatch "RUNSUBSCRIPT") {
  #run an external msl script
  $log.WriteLine("RUN SUB SCRIPT")
  handleRunSubScript($theLine)
 }
 elseif ($verb -imatch "RUN") {
  #run an external program
  $log.WriteLine("RUN")
  handleRUN($theLine)
 }
 elseif ($verb -imatch "WAIT") {
  #run an external program
  $log.WriteLine("WAIT")
  handleWaitSeconds($theLine)
 }
 elseif ($verb -imatch "WRITELN") {
  #writeLine to host
  $log.WriteLine("WRITELN")
  handleWriteLine($theLine)
 }
 elseif ($verb -imatch "SETMOUSE") {
  #run an external program
  $log.WriteLine("SET MOUSE")
  handleSetMousePt($theLine)
 }
 elseif ($verb -imatch "LEFTCLICKMOUSE") {
  #left click at the pt previously set by SETMOUSE
  $log.WriteLine("LEFT CLICK MOUSE")
  handleLeftClick($theLine)
 }
 elseif ($verb -imatch "SENDKEYS") {
  #left click at the pt previously set by SETMOUSE
  $log.WriteLine("SEND KEYS")
  handleSendKeys($theLine)
 }
 elseif ($verb -imatch "CLOSEWINDOWNAME") {
  #close the process with the given name
  #NB: it will close ALL instances with the given name
  $log.WriteLine("CLOSE WINDOW NAME")
  handleCloseWindowName($theLine)
 }
 else {  #unrecognised command
  #fail quietly for now
  $log.WriteLine("Error on Ln "+$linenumber+": verb="+$verb+" is unrecognised/unimplemented.")
  Write-Host ("Error on Ln "+$linenumber+": verb="+$verb+" is unrecognised/unimplemented.")
 }
 $lineNumber = $lineNumber + 1       #move on to the next line of the users script

}

$log.WriteLine("The script has ended. $(Get-Date).")
$log.Close()


