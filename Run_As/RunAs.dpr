program RunAs;

{$APPTYPE CONSOLE}

uses
  SysUtils,Windows;

function CreateProcessWithLogon(lpUsername        :PWideChar;
                                lpDomain          :PWideChar;
                                lpPassword        :PWideChar;
                                dwLogonFlags      :DWORD;
                                lpApplicationName :PWideChar;
                                lpCommandLine     :PWideChar;
                                dwCreationFlags   :DWORD;
                                lpEnvironment     :Pointer;
                                lpCurrentDirectory:PWideChar;
                                var lpStartupInfo :TStartupInfo;
                                var lpProcessInfo :TProcessInformation):BOOL;stdcall;external 'advapi32' name 'CreateProcessWithLogonW';

function CreateEnvironmentBlock(var lpEnvironment:Pointer;hToken:THandle;bInherit:BOOL):BOOL;stdcall;external 'userenv';
function DestroyEnvironmentBlock(pEnvironment:Pointer):BOOL;stdcall;external 'userenv';

const
  LOGON_WITH_PROFILE=$00000001;

procedure Error(s:string);
begin
  raise Exception.Create(s);
end;

procedure OSError(s:string);
(*
  Raise the last system error with an additional prefix message
*)
begin
  raise Exception.Create(s+#13#10+SysErrorMessage(GetLastError));
end;

function FormatParam(s:string):string;
(*
  Enclose into quotes (if not already) the string if it contains white space and return it, otherwise return the string itself.
*)
var
  a:Integer;
  t:Boolean;
begin
  Result:=s;
  t:=False;
  for a:=Length(s) downto 1 do
    if s[a] in [' ',#32] then begin
      t:=True;
      Break;
    end;
  if t and not (s[1] in ['''','"']) then
    Result:='"'+s+'"';
end;

function RunProcessAs(Command:string;Parameters:array of string;Username,Password:string;Domain:string='';WorkingDirectory:string='';Wait:Boolean=False):Cardinal;
(*
  Execute the Command with the given Parameters, Username, Domain, Password and Working Directory. Parameters containing white spaces are
  automatically embraced into quotes before being sent to avoid having them splitted by the system. If either Domain or Working Directory
  are empty the current one will be used instead.

  If Wait is specified the function will wait till the command is completely executed and will return the exit code of the process,
  otherwise zero.

  Suitable Delphi exceptions will be thrown in case of API failure.
*)
var
  a:Integer;
  n:Cardinal;
  h:THandle;
  p:Pointer;
  PI:TProcessInformation;
  SI:TStartupInfo;
  t:array[0..MAX_PATH] of WideChar;
  wUser,wDomain,wPassword,wCommandLine,wCurrentDirectory:WideString;
begin
  ZeroMemory(@PI,SizeOf(PI));
  ZeroMemory(@SI,SizeOf(SI));
  SI.cb:=SizeOf(SI);
  if not LogonUser(PChar(Username),nil,PChar(Password),LOGON32_LOGON_INTERACTIVE,LOGON32_PROVIDER_DEFAULT,h) then
    OSError('Could not log user in');
  try
    if not CreateEnvironmentBlock(p,h,True) then
      OSError('Could not access user environment');
    try
      wUser:=Username;
      wPassword:=Password;
      wCommandLine:=Command;
      for a:=Low(Parameters) to High(Parameters) do
        wCommandLine:=wCommandLine+' '+FormatParam(Parameters[a]);
      if Domain='' then begin
        n:=SizeOf(t);
        if not GetComputerNameW(t,n) then
          OSError('Could not get computer name');
        wDomain:=t;
      end else
        wDomain:=Domain;
      if WorkingDirectory='' then
        wCurrentDirectory:=GetCurrentDir
      else
        wCurrentDirectory:=WorkingDirectory;
      if not CreateProcessWithLogon(PWideChar(wUser),PWideChar(wDomain),PWideChar(wPassword),LOGON_WITH_PROFILE,nil,PWideChar(wCommandLine),CREATE_UNICODE_ENVIRONMENT,p,PWideChar(wCurrentDirectory),SI,PI) then
        OSError('Could not create process');
      if Wait then begin
        WaitForSingleObject(PI.hProcess,INFINITE);
        if not GetExitCodeProcess(PI.hProcess,Result) then
          OSError('Could not get process exit code');
      end else
        Result:=0;
      CloseHandle(PI.hProcess);
      CloseHandle(PI.hThread);
    finally
      DestroyEnvironmentBlock(p);
    end;
  finally
    CloseHandle(h);
  end;
end;

function FindStr(s:string;t:array of string):Integer;
(*
  Return the (case-insensitive) index of s into the array t, otherwise -1
*)
var
  a:Integer;
begin
  Result:=-1;
  for a:=Low(t) to High(T) do
    if AnsiUpperCase(t[a])=AnsiUpperCase(s) then begin
      Result:=a;
      Exit;
    end;
end;

function WaitChar:Char;
(*
  Wait till a character is typed in the console, and return its value
*)
var
  h:THandle;
  n:Cardinal;
  r:TInputRecord;
begin
  h:=GetStdHandle(STD_INPUT_HANDLE);
  repeat
    ReadConsoleInput(h,r,1,n);
  until (n=1) and (r.EventType=KEY_EVENT) and (r.Event.KeyEvent.bKeyDown);
  Result:=r.Event.KeyEvent.AsciiChar;
end;

function ReadlnMasked(var s:string):Boolean;
(*
  Read a string from the console input with masked characters, till either ENTER or ESCAPE is given. Return True if that was ENTER.
*)
var
  c:Char;
const
  EndChars:set of char=[#10,#13,#27];
begin
  s:='';
  repeat
    c:=WaitChar;
    if not (c in EndChars) then begin
      s:=s+c;
      Write('*');
    end;
  until c in EndChars;
  Result:=c=#13;
  WriteLn;
end;

procedure Main;
(*
  Main program, parse and process arguments, print usage if no arguments, ask for username or password if not specified and launch the
  desired process 
*)
var
  a,b:Integer;
  t:array of string;
  Username,Password,Domain,WorkingDir:string;

  function ExtractNext:string;
  (*
    Extract the next command-line argument, in respect of the previous flag
  *)
  begin
    Inc(b);
    if b>ParamCount then
      raise Exception.Create('Missing argument after '+ParamStr(b-1));
    Result:=ParamStr(b);
  end;

  procedure Parse;
  (*
    Parse the command-line flags
  *)
  var
    t:Boolean;
  begin
    t:=True;
    while t and (b<=ParamCount) do begin
      case FindStr(ParamStr(b),['---U','---P','---D','---W']) of
        0:Username:=ExtractNext;
        1:Password:=ExtractNext;
        2:Domain:=ExtractNext;
        3:WorkingDir:=ExtractNext;
      else
        t:=False;
      end;
      if t then
        Inc(b);
    end;
    if b>ParamCount then
      raise Exception.Create('Missing command name');
  end;

begin
  if ParamCount=0 then begin
    WriteLn('Usage: ',ChangeFileExt(ExtractFileName(ParamStr(0)),''),' [---U Username] [---D Domain] [---P Password] [---W WorkingDirectory] Command [Params]');
    ExitCode:=0;
    Exit;
  end;
  b:=1;
  Username:='';
  Password:='';
  Domain:='';
  WorkingDir:='';
  Parse;
  if Username='' then begin
    Write('Username: ');
    ReadLn(Username);
  end;
  if Password='' then begin
    Write('Password for ',Username,': ');
    if not ReadLnMasked(Password) then
      Error('Aborted');
  end;
  SetLength(t,ParamCount-b);
  try
    for a:=b+1 to ParamCount do
      t[a-b-1]:=ParamStr(a);
    ExitCode:=RunProcessAs(ParamStr(b),t,Username,Password,Domain,WorkingDir);
  finally
    SetLength(t,0);
  end;
end;

begin
  ExitCode:=-1;
  try
    Main;
  except
    on e:Exception do begin
      WriteLn('Error: ',e.Message);
      WriteLn('(Press any key to continue)');
      WaitChar;
    end;
  end;
end.