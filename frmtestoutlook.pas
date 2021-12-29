unit frmtestoutlook;

interface

uses
    Windows
  , Messages
  , SysUtils
  , Variants
  , Classes
  , Graphics
  , Controls
  , Forms
  , Dialogs
  , StdCtrls
  , System.Win.Registry
  , System.IniFiles
  , FireDAC.Stan.Intf
  , FireDAC.Stan.Option
  , FireDAC.Stan.Error
  , FireDAC.UI.Intf
  , FireDAC.Phys.Intf
  , FireDAC.Stan.Def
  , FireDAC.Stan.Pool
  , FireDAC.Stan.Async
  , FireDAC.Phys
  , FireDAC.Phys.MySQL
  , FireDAC.Phys.MySQLDef
  , FireDAC.VCLUI.Wait
  , Data.DB
  , FireDAC.Comp.Client
  , FireDAC.Stan.Param
  , FireDAC.DatS
  , FireDAC.DApt.Intf
  , FireDAC.DApt
  , FireDAC.Comp.DataSet
  , EmailSettings
  ;

type
  TForm2 = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    FDConnection1: TFDConnection;
    FDPhysMySQLDriverLink1: TFDPhysMySQLDriverLink;
    FDTable1: TFDTable;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    FIniSettings : TIniFile;
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

uses
    System.IOUtils
  ;


function DecodeProductKey(const HexSrc: array of Byte): string;
const
  StartOffset: Integer = $34; { //Offset 34 = Array[52] }
  EndOffset: Integer   = $34 + 15; { //Offset 34 + 15(Bytes) = Array[64] }
  Digits: array[0..23] of CHAR = ('B', 'C', 'D', 'F', 'G', 'H', 'J',
    'K', 'M', 'P', 'Q', 'R', 'T', 'V', 'W', 'X', 'Y', '2', '3', '4', '6', '7', '8', '9');
  dLen: Integer = 29; { //Length of Decoded Product Key }
  sLen: Integer = 15;
  { //Length of Encoded Product Key in Bytes (An total of 30 in chars) }
var
  HexDigitalPID: array of CARDINAL;
  Des: array of CHAR;
  I, N: INTEGER;
  HN, Value: CARDINAL;
begin
  SetLength(HexDigitalPID, dLen);
  for I := StartOffset to EndOffset do
  begin
    HexDigitalPID[I - StartOffSet] := HexSrc[I];
  end;

  SetLength(Des, dLen + 1);

  for I := dLen - 1 downto 0 do
  begin
    if (((I + 1) mod 6) = 0) then
    begin
      Des[I] := '-';
    end
    else
    begin
      HN := 0;
      for N := sLen - 1 downto 0 do
      begin
        Value := (HN shl 8) or HexDigitalPID[N];
        HexDigitalPID[N] := Value div 24;
        HN    := Value mod 24;
      end;
      Des[I] := Digits[HN];
    end;
  end;
  Des[dLen] := Chr(0);

  for I := 0 to Length(Des) do
  begin
    Result := Result + Des[I];
  end;
end;

procedure TForm2.Button1Click(Sender: TObject);
var
  email : TEmailSettings;
begin
  email := TEmailSettings.Create;
  Memo1.Lines.Text := email.Debug;
end;

procedure TForm2.Button2Click(Sender: TObject);
var
  profiles : TOutlookProfiles;
  i : Integer;
  j , k : Integer;
begin
 profiles := TOutlookProfiles.Create;
 try
   Memo1.Lines.Add('ProfileCount:' + IntToStr(profiles.Count));
    for i := 0 to profiles.Count-1 do
    begin
      Memo1.Lines.Add('ProfileName:' + profiles[i].ProfileName);
      Memo1.Lines.Add('AccountCount:' + IntToStr(profiles[i].Count));
      for j := 0 to profiles[i].Count - 1 do
      begin
        Memo1.Lines.Add('AccountName:' + profiles[i].Accounts[j].DisplayName);
        Memo1.Lines.Add('AccountDataPath:' + profiles[i].Accounts[j].DataFilePath);
        if profiles[i].Accounts[j].IsAddressBook then
          continue;


        Memo1.Lines.Add('AccountPOP3:' + profiles[i].Accounts[j].POP3Server);
        Memo1.Lines.Add('AccountIMAPMailServer:' + profiles[i].Accounts[j].IMAPMailServer);
        Memo1.Lines.Add('AccountSMTP:' + profiles[i].Accounts[j].SMTPMailServer);
        Memo1.Lines.Add('AccountSUID:' + profiles[i].Accounts[j].ServiceUID);

        Memo1.Lines.Add('AccountReg:' + profiles[i].Accounts[j].RegistryKey);
        Memo1.Lines.Add('AccountPUID:' + profiles[i].Accounts[j].PreferenceUID);



        for k := 0 to profiles[i].Accounts[j].GetSafeSendersCount()-1 do
        begin
          Memo1.Lines.Add('Safe['+IntToStr(k)+']:'+ profiles[i].Accounts[j].SafeSenders[k]);
        end;
        for k := 0 to profiles[i].Accounts[j].GetSafeRecipientCount()-1 do
        begin
          Memo1.Lines.Add('SafeRecipient['+IntToStr(k)+']:'+ profiles[i].Accounts[j].SafeRecipients[k]);
        end;


        for k := 0 to profiles[i].Accounts[j].GetBlockedSenderCount()-1 do
        begin
          Memo1.Lines.Add('BlockedSenders['+IntToStr(k)+']:'+ profiles[i].Accounts[j].BlockedSenders[k]);
          Application.ProcessMessages;
          FDTable1.Active := True;
          if not FDTable1.Locate('source', profiles[i].Accounts[j].BlockedSenders[k], []) then
          begin
            FDTable1.Append;
            try
              Memo1.Lines.Add('AddedBlockedSenders['+IntToStr(k)+']:'+ profiles[i].Accounts[j].BlockedSenders[k]);
              FDTable1.FieldByName('source').AsString := profiles[i].Accounts[j].BlockedSenders[k];
              FDTable1.FieldByName('sys_userid').AsInteger := 1;
              FDTable1.FieldByName('sys_groupid').AsInteger := 1;
              FDTable1.FieldByName('sys_perm_user').AsString := 'riud';
              FDTable1.FieldByName('sys_perm_group').AsString := 'riud';
              FDTable1.FieldByName('server_id').AsInteger := 1;
              FDTable1.FieldByName('access').AsString := 'REJECT';
              FDTable1.FieldByName('type').AsString := 'sender';
              FDTable1.FieldByName('active').AsString := 'y';
              FDTable1.Post;
            except
              FDTable1.Cancel;
            end;
          end;
        end;
        //001f0418
      end;
    end;
 finally
   FreeAndNil(profiles);
 end;
end;

procedure TForm2.Button3Click(Sender: TObject);
var
  Profiles : TWindowsProfiles;
  i : Integer;
begin
  Profiles := TWindowsProfiles.Create;
  try
   for i  := 0 to profiles.count-1 do
   begin
     Memo1.Lines.Add('ProfilePath:'+profiles[i].ProfilePath);
   end;
  finally
    FreeAndNil(profiles);
  end;
end;

procedure Priv();
var
  TTokenHd: THandle;
  TTokenPvg: TTokenPrivileges;
  cbtpPrevious: DWORD;
  rTTokenPvg: TTokenPrivileges;
  pcbtpPreviousRequired: DWORD;
  tpResult: Boolean;
const
  SE_SHUTDOWN_NAME = 'SeShutdownPrivilege';
  SE_BACKUP_NAME   = 'SeBackupPrivilege';
  SE_RESTORE_NAME  = 'SeRestorePrivilege';
begin
   if Win32Platform = VER_PLATFORM_WIN32_NT then
   begin
     tpResult := OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES or TOKEN_QUERY, TTokenHd);
     if tpResult then
     begin
       tpResult := LookupPrivilegeValue(nil,
                                        SE_BACKUP_NAME,
                                        TTokenPvg.Privileges[0].Luid);
       TTokenPvg.PrivilegeCount := 1;
       TTokenPvg.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED;
       cbtpPrevious := SizeOf(rTTokenPvg);
       pcbtpPreviousRequired := 0;
       if tpResult then
         Windows.AdjustTokenPrivileges(TTokenHd,
                                       False,
                                       TTokenPvg,
                                       cbtpPrevious,
                                       rTTokenPvg,
                                       pcbtpPreviousRequired);

       tpResult := LookupPrivilegeValue(nil,
                                        SE_RESTORE_NAME,
                                        TTokenPvg.Privileges[0].Luid);
       TTokenPvg.PrivilegeCount := 1;
       TTokenPvg.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED;
       cbtpPrevious := SizeOf(rTTokenPvg);
       pcbtpPreviousRequired := 0;
       if tpResult then
         Windows.AdjustTokenPrivileges(TTokenHd,
                                       False,
                                       TTokenPvg,
                                       cbtpPrevious,
                                       rTTokenPvg,
                                       pcbtpPreviousRequired);
     end;
   end;
end;

procedure TForm2.Button4Click(Sender: TObject);
begin
  Priv();
  LoadUserHive('C:\Users\geoff\Desktop\config\SOFTWARE');
end;

procedure TForm2.Button5Click(Sender: TObject);
var
  reg : TRegistry;
  str : String;
begin
  reg := TRegistry.Create;
  try
    reg.RootKey := HKEY_LOCAL_MACHINE;
    reg.OpenKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion', False);
    returnBinaryValueAsString(reg,'DigitalProductID', str);
    Memo1.Lines.Add('result:' + str);
//    Memo1.Lines.Add(GetProductKey);
  finally
    FreeAndNil(reg);
  end;
end;

function GetCurrentUserAndDomain1(szUser : PChar;   pcchUser : DWORD;
       szDomain : PChar;  pcchDomain: DWORD) : boolean;
var
  fSuccess : boolean;
  hToken   : THandle;
  ptiUser  : PSIDAndAttributes;
  cbti     : DWORD;
  snu      : SID_NAME_USE;
begin
  ptiUser := nil;
  Result := false;

  try
    // Get the calling thread's access token.
    if (not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, TRUE, hToken)) then
     begin
       if (GetLastError() <> ERROR_NO_TOKEN) then
         Exit;

       // Retry against process token if no thread token exists.
       if (not OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hToken)) then
         Exit;
     end;

    // Obtain the size of the user information in the token.
    if (GetTokenInformation(hToken, TokenUser, nil, 0, cbti)) then
      Exit // Call should have failed due to zero-length buffer.
    else if (GetLastError() <> ERROR_INSUFFICIENT_BUFFER) then
      Exit; // Call should have failed due to zero-length buffer.

    // Allocate buffer for user information in the token.
    ptiUser :=  HeapAlloc(GetProcessHeap(), 0, cbti);
    if (ptiUser= nil) then
      Exit;

    // Retrieve the user information from the token.
    if ( not GetTokenInformation(hToken, TokenUser, ptiUser, cbti, cbti)) then
      Exit;

    // Retrieve user name and domain name based on user's SID.
    if ( not LookupAccountSid(nil, ptiUser.Sid, szUser, pcchUser, szDomain, pcchDomain, snu)) then
      Exit;

    fSuccess := TRUE;
  finally
    // Free resources.
    if (hToken > 0) then
      CloseHandle(hToken);

    if (ptiUser <> nil) then
      HeapFree(GetProcessHeap(), 0, ptiUser);
  end;
  Result :=  fSuccess;
end;

function GetCurrentUserAndDomain(var susername:String; var sdomain:String) : boolean;
var
  user : array [0..200] of Char;
  domain: array [0..200] of Char;
begin
  Result := GetCurrentUserAndDomain1(user, length(user), domain, length(domain));
  susername := user;
  sdomain   := domain;
end;

procedure TForm2.Button6Click(Sender: TObject);
var
  user : String;
  domain: String;
begin
  //GetCurrentUserAndDomain1(user,length(user),domain,length(domain));
  GetCurrentUserAndDomain(user, domain);
  Memo1.Lines.Add('username:' + user);
  Memo1.Lines.Add('domain:' + domain);
end;



procedure TForm2.FormCreate(Sender: TObject);
var
  filename : string;
  oParams: TFDPhysMySQLConnectionDefParams;
begin
  FDPhysMySQLDriverLink1.VendorLib := TPath.Combine(TPath.GetDirectoryName(Application.ExeName), 'libmysql.dll');
  filename := TPath.Combine(TPath.GetDirectoryName(Application.ExeName), 'email.ini');
  FIniSettings := TIniFile.Create(filename);
  oParams := FDConnection1.Params as TFDPhysMySQLConnectionDefParams;
  oParams.UserName := FIniSettings.ReadString('Authentication', 'Username', '');
  oParams.Password := FIniSettings.ReadString('Authentication', 'Password', '');
  oParams.Database := FIniSettings.ReadString('Authentication', 'Database', '');
  oParams.Server   := FIniSettings.ReadString('Authentication', 'Server', '');
end;

procedure TForm2.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FIniSettings);
end;

end.
