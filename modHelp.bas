Attribute VB_Name = "modHelp"
Option Explicit

Sub Assign_Help()

strHelp = strProgrammeName & " " & strVersion & vbCrLf & vbCrLf

strHelp = strHelp & _
            strProgrammeName & " is (domain) name server of sorts." & _
            " It listens on UDP port 53 (by default) for DNS" & vbCrLf & _
            "requests and replies to those requests if it has the" & _
            " necessary information. " & strProgrammeName & " " & vbCrLf & _
            "understands A, AAAA" & _
            " and PTR requests." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "http://www.cyodesigns.com/hoosns/" & vbCrLf & vbCrLf
            
strHelp = strHelp & vbTab & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf

strHelp = strHelp & "'" & LCase(strProgrammeName) & ".exe' and '" & _
            hostFN & "'" & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            strProgrammeName & " retrieves " & _
            "name/IP mappings from the equivalent of a hosts file, called '" & hostFN & "'." & _
            "  " & vbCrLf & _
            "Entries in this file follow the same format as a typical" & _
            " hosts file:" & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "192.168.1.254" & vbTab & "gateway" & vbCrLf & _
            "192.168.1.2" & vbTab & "server" & vbCrLf & vbCrLf

strHelp = strHelp & _
            "All that is required to start providing name resolution services " & _
            "is the " & strProgrammeName & " 'exe' and some" & vbCrLf & _
            "valid entries in a '" & hostFN & "' file located in the same" & _
            " directory as the 'exe'. " & strProgrammeName & " parses the " & vbCrLf & _
            "'" & hostFN & "' when it loads. Any lines containing valid" & _
            " name to IP mappings are displayed under" & vbCrLf & _
            "the 'Name/IP Mapping' radio button. Changes can be made to '" & hostFN & "'" & _
            " while " & strProgrammeName & " is " & vbCrLf & _
            "running. Press the 'Reload' button" & _
            " to refresh" & vbCrLf & vbCrLf

strHelp = strHelp & vbTab & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf

strHelp = strHelp & "'" & iniFile & "'" & vbCrLf & vbCrLf

strHelp = strHelp & _
            strProgrammeName & " behaviour can be modified by changing the contents" & _
            " of an ini file called '" & iniFile & "'." & vbCrLf & _
            "This file must also reside in the same directory as the 'exe'" & _
            " when it loads. Options include:" & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "port = 53" & vbCrLf & _
            "Determines which UDP port " & strProgrammeName & " listens for requests." & vbCrLf & vbCrLf

strHelp = strHelp & _
            "createptr = 1" & vbCrLf & _
            "A PTR record can be created for every valid Name/IP entry in '" & _
            hostFN & ". '1' will create PTR" & vbCrLf & "records, '0' will not.  " & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "listeningip = 192.168.1.53" & vbCrLf & _
            "On a machine with more than one assigned IP, this setting allows " & _
            "only one IP to listen instead" & vbCrLf & _
            "of all IPs." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "timetolive = 3600" & vbCrLf & _
            "DNS replies need a ttl (time to live) value. This value is in seconds and applies to" & _
            " all replies." & vbCrLf & vbCrLf

strHelp = strHelp & _
            "logactivitytofile = 0" & vbCrLf & _
            "By default " & strProgrammeName & " will display up to 100 of the" & _
            " most recent replies under the 'Activity'" & vbCrLf & _
            "radio button. This option " & _
            "instructs " & strProgrammeName & " to log activity to a text file called '" & _
            logRepliesFile & "'." & vbCrLf & _
            "'1' switches logging on, '0' switches it off." & vbCrLf & vbCrLf
            

strHelp = strHelp & _
            "guihide = 0" & vbCrLf & _
            "Allows " & strProgrammeName & " to execute without a GUI." & _
            " At the moment the only way to terminate the" & vbCrLf & _
               strProgrammeName & " process when running without a GUI is to" & _
            " use taskmgr.exe or similar." & _
            "'1' hides the GUI," & vbCrLf & "'0' displays it." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "startpaused = 0" & vbCrLf & _
             strProgrammeName & " will not respond to requests on start up if this is set to '1'." & _
             "" & vbCrLf & vbCrLf

strHelp = strHelp & vbTab & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf
      
strHelp = strHelp & "'" & logFile & "'" & vbCrLf & vbCrLf

strHelp = strHelp & _
            "By default " & strProgrammeName & " will log events such as starting" & _
            " and stopping in a file called '" & logFile & "'." & vbCrLf & _
            "Also, invalid lines that it finds in '" & hostFN & "' and '" & iniFile & "'" & _
            " are logged to this log file." & vbCrLf & " The '" & logFile & "' is over written every " & _
            "time  " & strProgrammeName & " starts." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "This logging behaviour can be stopped by passing 'nosystemlog' as a command" & vbCrLf & _
            "line parameter. Note: this parameter doesn't affect DNS activity logging." & vbCrLf & vbCrLf
            
strHelp = strHelp & vbTab & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf
            
strHelp = strHelp & "'Name/IP Mapping' Radio Button" & vbCrLf & vbCrLf

strHelp = strHelp & _
            "Name/IP mapping lines that " & strProgrammeName & " finds in '" & hostFN & "'" & _
            " are found in the " & _
            "'Name/IP Mapping' " & vbCrLf & _
            "page. Each line will look similar to this:" & vbCrLf & vbCrLf & _
            "0   192.168.0.10    server" & vbCrLf & vbCrLf & _
            "The first column contains a number, this indicates how many times the name to the right has" & vbCrLf & _
            "been requested. When the request is made " & strProgrammeName & " places a" & _
            " '<' character to the right of the " & vbCrLf & _
            "cumulative total:" & vbCrLf & vbCrLf & _
            "1<   192.168.0.10    server" & vbCrLf & vbCrLf & _
            "Therefore, at most one line will have a '<' on it. This will" & _
            " indicate the most recent successful" & vbCrLf & _
            "request." & vbCrLf & vbCrLf

strHelp = strHelp & vbTab & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf
            
strHelp = strHelp & "'Activity' Radio Button" & vbCrLf & vbCrLf

strHelp = strHelp & _
            "Lines on the Activity radio button show " & strProgrammeName & " activity." & _
            " Each line corresponds to a" & vbCrLf & _
            "request/response event between " & strProgrammeName & _
            " and a requesting host. Each line looks similar to this:" & vbCrLf & vbCrLf
            
strHelp = strHelp & "+ A  17/04/2008 19:10:35 server -- 192.168.1.78" & vbCrLf & vbCrLf
         
strHelp = strHelp & _
            "The '+' character shows that the activity has successful and that a reply was sent." & _
            " A " & vbCrLf & "'.' character indicates that " & strProgrammeName & _
            " didn't have a name mapping for the request." & _
            " A '?' indicates " & vbCrLf & _
            "that the query type was unknown. has success full and that a reply has been sent." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "The next character indicates the query type. In this example the 'A' character." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "The next two columns are the date and time of the activity." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "The '--' characters are delimiters." & vbCrLf & vbCrLf
            
strHelp = strHelp & _
            "The final column indicates the requesting IP." & vbCrLf & vbCrLf
            
strHelp = strHelp & vbTab & "------------------------------------------------------------------------------------" & vbCrLf & vbCrLf
            
strHelp = strHelp & "THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND," & vbCrLf & _
            "EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF" & vbCrLf & _
            "MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND" & vbCrLf & _
            "NONINFRINGEMENT.  IN NO EVENT SHALL THE COPYRIGHT HOLDERS BE LIABLE" & vbCrLf & _
            "FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF" & vbCrLf & _
            "CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION" & vbCrLf & _
            "WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE." & vbCrLf & vbCrLf
    
End Sub

