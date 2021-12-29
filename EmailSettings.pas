unit EmailSettings;

interface

uses
    Registry
  , Sysutils
  , Classes
  , Windows
  , Generics.Collections
  ;

type
  TBaseAccount = class
    private
     reg              : TRegistry;
     SSafeSenders     : String;
     FSafeSenders     : TStringList;
     FSafeRecipients  : TStringList;
     FBlockedSenders  : TStringList;

     function GetSafeSenders(i:Integer): String;
     procedure SetSafeSenders(i:Integer; const x : String);
     function GetSafeRecipients(i:Integer): String;
     procedure SetSafeRecipients(i:Integer; const x : String);
     function GetBlockedSenders(i:Integer): String;
     procedure SetBlockedSenders(i:Integer; const x : String);

    public
     DisplayName      : String;
     SMTPMailServer   : String;
     SMTPPort         : Integer;
     SMTPUseSSL       : Boolean;
     IMAPMailServer   : String;
     IMAPPort         : Integer;
     EmailAddress     : String;
     POP3Server       : String;
     POP3Port         : Integer;
     UseSSLOnRetrieve : Boolean;
     RegistryKey      : String;
     AccessProgram    : String;
     DataFilePath     : String;
     ServiceUID       : String;
     ServiceUIDnew    : String;
     PreferenceUID    : String;
     constructor Create(regKey: String);
     destructor Destroy(); override;

     ///	<summary>
     ///	  Is used to determine if the account has a local store.
     ///	</summary>
     ///	<returns>
     ///	  Returns TRUE if the account is/has a local store.
     ///	</returns>
     function isStore(): Boolean;

     ///	<summary>
     ///	  Is used to determine if the account is/has a Address Book.
     ///	</summary>
     ///	<returns>
     ///	  Returns TRUE if the account is/has an Address Book.
     ///	</returns>
     function IsAddressBook(): Boolean;

     ///	<summary>
     ///	  Is used to determine if the account is a mailbox.
     ///	</summary>
     ///	<returns>
     ///	  Returns TRUE if the account is a mailbox.
     ///	</returns>
     function IsMailbox(): Boolean;
     function GetSafeSendersCount(): Integer;
     function GetBlockedSenderCount(): Integer;

     ///	<summary>
     ///	  Returns the number of Blocked Sender email addresses
     ///	</summary>
     function GetSafeRecipientCount(): Integer;
     property SafeSenders[i:Integer]: String read GetSafeSenders write SetSafeSenders;

     ///	<summary>
     ///	  Contains Recipients that have been deemed to be safe to receive
     ///	  mail from.
     ///	</summary>
     ///	<param name="i">
     ///	  Integer representing position in array
     ///	</param>
     property SafeRecipients[i:Integer]: String read GetSafeRecipients write SetSafeRecipients;
     property BlockedSenders[i:Integer]: String read GetBlockedSenders write SetBlockedSenders;

  end;

  TSMTPAccount = class (TBaseAccount)

  end;

  TPOP3SMTPAccount = class(TSMTPAccount)

  end;

  TIMAPSMTPAccount = class(TSMTPAccount)

  end;

  TExchangeAccount = class(TBaseAccount)

  end;

  ///	<summary>
  ///	  Contains all the information for a particular Outlook Profile.  This is
  ///	  mostly stored in a list of accounts.
  ///	</summary>
  TOutlookProfile = class
    private
      reg : TRegistry;
      Key : String;
      FAccounts : TObjectList<TBaseAccount>;
      function GetProfileAccount(i : Integer) : TBaseAccount;
    public
      ///	<summary>
      ///	  The name of the Profile.
      ///	</summary>
      ProfileName : String;
      constructor Create(hKey: String; inProfileName: String); reintroduce;

      ///	<returns>
      ///	  Returns the number of Accounts in the profile.
      ///	</returns>
      function Count():Integer;
      destructor Destroy; override;
      property Accounts[i : Integer] : TBaseAccount read GetProfileAccount; default;
  end;

  ///	<summary>
  ///	  Is a list of Profiles that are in the Microsoft Outlook
  ///	</summary>
  TOutlookProfiles = class
    private
      FProfiles : TObjectList<TOutlookProfile>;
      reg : TRegistry;
      function GetOutlookProfile(i : Integer) : TOutlookProfile;
    public
      function Count():Integer;
      constructor Create;
      destructor Destroy; override;

      ///	<summary>
      ///	  Contains an array of the Outlook profiles in the current windows
      ///	  user profile.
      ///	</summary>
      ///	<param name="i">
      ///	  index to the Outlook Profile.
      ///	</param>
      property Profiles[i : Integer] : TOutlookProfile read GetOutlookProfile; default;
  end;

  TUserPrivileges = class

  end;

  TWindowsUserProfile = class
    private
      reg : TRegistry;
      function GetProfilePath(): String;
    public
      constructor Create(key:String);
      destructor Destroy; override;
      property ProfilePath: String read GetProfilePath;
  end;

  TWindowsProfiles = class
    private
      reg : TRegistry;
      FAccounts : TObjectList<TWindowsUserProfile>;
      function GetAccountByIndex(i : Integer) : TWindowsUserProfile;
    public
      constructor Create;
      destructor Destroy; override;
      function Count(): Integer;
      property Accounts[i : Integer] : TWindowsUserProfile read GetAccountByIndex; default;
  end;

  TEmailSettings = class
    private
      outlooksubKeys   : TStringList;
      outlooksubKeys2  : TStringList;
      expresssubKeys   : TStringList;
      outlookAccounts  : TRegistry;
      outlookAccounts2 : TRegistry;
      expressAccounts  : TRegistry;
      emailDetails     : TList;
      function FillInDetails(): TBaseAccount;
    public
      constructor Create;
      destructor Destroy; override;
      function GetEmailCount(): Integer;
      function FindEmailAccount(addr: String): TBaseAccount;
//      procedure Debug(); overload;
      function Debug(): String; overload;
  end;

procedure LoadUserHive(sPathToUserHive: string);
function returnBinaryValueAsString(reg:TRegistry; name:String; var value:String):Boolean;

implementation

const crlf = #10;

//NOTE:   sPathToUserHive is the full path to the users "ntuser.dat" file.
//

procedure LoadUserHive(sPathToUserHive: string);
var
  MyReg: TRegistry;
//  UserPriv: TUserPrivileges;
begin
//  UserPriv := TUserPrivileges.Create;
  try
//    with UserPriv do
    begin
//      if HoldsPrivilege(SE_BACKUP_NAME) and HoldsPrivilege(SE_RESTORE_NAME) then
      begin
//        PrivilegeByName(SE_BACKUP_NAME).Enabled := True;
//        PrivilegeByName(SE_RESTORE_NAME).Enabled := True;

        MyReg := TRegistry.Create;
        try
          MyReg.RootKey := HKEY_LOCAL_MACHINE;
          MyReg.UnLoadKey('TEMP_HIVE'); //unload hive to ensure one is not already loaded

          if MyReg.LoadKey('TEMP_HIVE', sPathToUserHive) then
          begin
            //ShowMessage( 'Loaded' );
            MyReg.OpenKey('TEMP_HIVE', False);

//            if MyReg.OpenKey('TEMP_HIVE\Environment', True) then
            begin
              // --- Make changes *here* ---
              //
//              MyReg.WriteString('KEY_TO_WRITE', 'VALUE_TO_WRITE');
              //
              //
            end;

            //Alright, close it up
//            MyReg.CloseKey;
//            MyReg.UnLoadKey('TEMP_HIVE');
            //let's unload the hive since we are done with it
          end
          else
          begin
//            Memo1.Lines.Add('Error Loading: ' + sPathToUserHive);
          end;
        finally
          FreeAndNil(MyReg);
        end;

      end;
//      WriteLn('Required privilege not held');
    end;
  finally
//    FreeAndNil(UserPriv);
  end;
end;

constructor TWindowsUserProfile.Create(key: String);
begin
  reg := TRegistry.Create;
  reg.RootKey := HKEY_LOCAL_MACHINE;
  reg.OpenKey(key, False);
end;

function TWindowsUserProfile.GetProfilePath(): String;
begin
  Result := reg.ReadString('ProfileImagePath');
end;

destructor TWindowsUserProfile.Destroy;
begin
  FreeAndNil(reg);
end;

constructor TWindowsProfiles.Create;
var
  valNames  : TStringList;
  valueName : String;
  account   : TWindowsUserProfile;
  rtc       : Boolean;
begin
  FAccounts := nil;
  reg := nil;
  valNames := nil;
  try
    FAccounts   := TObjectList<TWindowsUserProfile>.Create;
    reg         := TRegistry.Create;
    valNames    := TStringList.Create;
    reg.RootKey := HKEY_LOCAL_MACHINE;
    rtc := reg.OpenKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList', False);
    reg.GetKeyNames(valNames);
    for valueName in valNames do
    begin
      account := TWindowsUserProfile.Create('SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' + valueName);
      FAccounts.Add(account);
    end;
  finally
    FreeAndNil(valNames);
  end;
end;

function TWindowsProfiles.GetAccountByIndex(i : Integer): TWindowsUserProfile;
begin
  Result := FAccounts[i];
end;

function TWindowsProfiles.Count(): Integer;
begin
  Result := FAccounts.Count;
end;

destructor TWindowsProfiles.Destroy;
begin
  FreeAndNil(reg);
  FreeAndNil(FAccounts);
end;

constructor TOutlookProfile.Create(hKey: String; inProfileName:String);
var
  accountList : TStringList;
  accountName : String;
  account     : TBaseAccount;
  i : Integer;
begin
  reg := TRegistry.Create;
  FAccounts := TObjectList<TBaseAccount>.Create;
  accountList := TStringList.Create;
  try
    reg.RootKey := HKEY_CURRENT_USER;
    reg.OpenKey(hKey + '\9375CFF0413111d3B88A00104B2A6676', False);
    Key := hKey;
    reg.GetKeyNames(accountList);
    ProfileName := inProfileName;
    for i := 0 to accountList.Count - 1 do
    begin
      accountName := accountList[i];
      account := TBaseAccount.Create(hKey + '\9375CFF0413111d3B88A00104B2A6676\' + accountName);
      FAccounts.Add(account);
    end;
  finally
    FreeAndNil(accountList);
  end;
end;

function TBaseAccount.GetSafeSenders(i:Integer):String;
begin
  Result := FSafeSenders[i];
end;

procedure TBaseAccount.SetSafeSenders(i:Integer;const x : String);
begin
  FSafeSenders[i] := x;
end;

function TBaseAccount.GetBlockedSenders(i:Integer):String;
begin
  Result := FBlockedSenders[i];
end;

procedure TBaseAccount.SetBlockedSenders(i:Integer;const x : String);
begin
  FBlockedSenders[i] := x;
end;


function TBaseAccount.GetSafeRecipients(i:Integer):String;
begin
  Result := FSafeSenders[i];
end;

procedure TBaseAccount.SetSafeRecipients(i:Integer;const x : String);
begin
  FSafeSenders[i] := x;
end;

destructor TOutlookProfile.Destroy;
begin
  FreeAndNil(reg);
  FreeAndNil(FAccounts);
end;

function TOutlookProfile.Count():Integer;
begin
  Result := FAccounts.Count;
end;

function TOutlookProfiles.Count():Integer;
begin
  Result := FProfiles.Count;
end;

function TOutlookProfile.GetProfileAccount(i : Integer) : TBaseAccount;
begin
  Result := FAccounts[i];
end;

function TOutlookProfiles.GetOutlookProfile(i : Integer) : TOutlookProfile;
begin
  Result := FProfiles[i];
end;

constructor TOutlookProfiles.Create;
var
  profileNames : TStringList;
  profileName  : String;
  profile      : TOutlookProfile;
begin
  FProfiles    := TObjectList<TOutlookProfile>.Create(True);
  reg          := TRegistry.Create;
  profileNames := TStringList.Create;
  try
    reg.RootKey  := HKEY_CURRENT_USER;
    reg.OpenKey('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\', False);
    reg.GetKeyNames(profileNames);
    for profileName in ProfileNames do
    begin
      profile := TOutlookProfile.Create('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\' + profileName, profileName);
      FProfiles.Add(profile);
    end;
  finally
    FreeAndNil(profileNames);
  end;
end;

destructor TOutlookProfiles.Destroy;
begin
  FreeAndNil(FProfiles);
  FreeAndNil(reg);
end;

function ConvertBinaryToString(value:pbyte; len:Integer):string;
var
  hex: String;
  i : Integer;
begin
  Result := '';
  i := 0;
  while i < len do
  begin
    hex := hex + IntToHex(byte(value[i]), 2);
    Inc(i);
  end;
  Result := lowercase(hex);
end;

function returnUIDKey(reg:TRegistry; name:String; var value:String):Boolean;
var
  data  : String;
  dSize : Integer;
begin
  value := '';
  if reg.ValueExists(name) then
  begin
    dSize := reg.GetDataSize(name);
    setLength(data, dSize + 1);
    dsize := reg.ReadBinaryData(name, PByte(data)^, dSize);
    value := ConvertBinaryToString(@Data[1], dSize );
    Result := True;
  end
  else
  begin
     Value  := '';
     Result := False;
  end;
end;

function returnBinaryValueAsString(reg:TRegistry; name:String; var value:String):Boolean;
var
  dSize : Integer;
  data  : String;
begin
  if reg.ValueExists(name) then
  begin
    dSize := reg.GetDataSize(name);
    setLength(data, dSize);
    reg.ReadBinaryData(name, PByte(data)^, dSize);
    value := data;
    Result := True;
  end
  else
  begin
    value  := '';
    Result := False;
  end;
end;

constructor TBaseAccount.Create(regKey:String);
var
  Data          : String;
  regService    : TRegistry;
  regPreference : TRegistry;
  regPSTPath    : TRegistry;
  regOSTPath    : TRegistry;
  regPSTPathStr : String;
  regOSTPathStr : String;
begin
  RegistryKey                 := regKey;
  reg                         := TRegistry.Create;
  regService                  := TRegistry.Create;
  regPreference               := TRegistry.Create;
  regOSTPath                  := TRegistry.Create;
  regPSTPath                  := TRegistry.Create;


  try
    reg.RootKey               := HKEY_CURRENT_USER;
    regService.RootKey        := HKEY_CURRENT_USER;
    regPreference.RootKey     := HKEY_CURRENT_USER;
    regPSTPath.RootKey        := HKEY_CURRENT_USER;
    regOSTPath.RootKey        := HKEY_CURRENT_USER;

    FSafeSenders              := TStringList.Create;
    FSafeSenders.LineBreak    := ';';
    FBlockedSenders           := TStringList.Create;
    FBlockedSenders.LineBreak := ';';
    FSafeRecipients           := TStringList.Create;
    FSafeRecipients.LineBreak := ';';
    reg.OpenKey(RegistryKey,False);
    DisplayName               := '';
    SMTPMailServer            := '';
    SMTPPort                  := 0;
    EmailAddress              := '';
    POP3Server                := '';
    POP3Port                  := 0;
    DataFilePath              := '';
    AccessProgram             := 'Outlook';

    returnBinaryValueAsString(reg, 'POP3 Server', POP3Server);
    returnBinaryValueAsString(reg, 'IMAP Server', IMAPMailServer);
    returnBinaryValueAsString(reg, 'SMTP Server', SMTPMailServer);
    returnBinaryValueAsString(reg, 'Email', EmailAddress);
    returnBinaryValueAsString(reg, 'Account Name', DisplayName);

    if reg.ValueExists('SMTP Port') then
       SMTPPort    := reg.ReadInteger('SMTP Port')
    else SMTPPort  := 25;

    if reg.ValueExists('SMTP Use SSL') then
       SMTPUseSSL    := reg.ReadBool('SMTP Use SSL')
    else SMTPUseSSL  := False;

    if reg.ValueExists('SMTP Email Address') then
       EmailAddress   := reg.ReadString('SMTP Email Address')
    else EmailAddress := '';

    UseSSLOnRetrieve := False;

    if reg.ValueExists('POP3 Port') then
    begin
      POP3Port         := reg.ReadInteger('POP3 Port');
      if reg.ValueExists('POP3 Use SSL') then
        UseSSLOnRetrieve := reg.ReadBool('POP3 Use SSL');
    end
    else POP3Port := 110;

    if reg.ValueExists('IMAP Port') then
    begin
      IMAPPort         := reg.ReadInteger('IMAP Port');
      if reg.ValueExists('IMAP Use SSL') then
        UseSSLOnRetrieve := reg.ReadBool('IMAP Use SSL');
    end
    else IMAPPort := 143;

    if returnUIDKey(reg, 'Preferences UID', PreferenceUID) then  // try and locate data file
    begin
      if regService.OpenKey('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\' + PreferenceUID, False) then
      begin
        returnBinaryValueAsString(regService, '001f0418', Data);
        FSafeSenders.Text := Data;
        returnBinaryValueAsString(regService, '001f0419', Data);
        FSafeRecipients.Text := Data;
        returnBinaryValueAsString(regService, '001f041a', Data);
        FBlockedSenders.Text := Data.Trim;
      end;
    end;

    if returnUIDKey(reg, 'Service UID', ServiceUID) then  // try and locate data file
    begin
      if regPreference.OpenKey('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\' + ServiceUID, False) then
      begin
        returnUIDKey(regPreference, '01023d00', regPSTPathStr);
        returnUIDKey(regPreference, '01023d15', regOSTPathStr);
        if((Length(regPSTPathStr) > 0) and regPSTPath.OpenKey('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\' + regPSTPathStr, False)) then
        begin
          returnBinaryValueAsString(regPSTPath, '001f6700', DataFilePath);
        end;
        if Length(DataFilePath) = 0 then
        begin
          if((Length(regOSTPathStr) > 0) and regOSTPath.OpenKey('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\' + regOSTPathStr, False)) then
          begin
            returnBinaryValueAsString(regOSTPath, '001f6610', DataFilePath);
          end;
        end;
      end;
    end;
  finally
    FreeAndNil(regService);
    FreeAndNil(regPreference);
    FreeAndNil(regOSTPath);
    FreeAndNil(regPSTPath);
  end;
end;

function TBaseAccount.GetSafeSendersCount():Integer;
begin
  Result := FSafeSenders.Count;
end;

function TBaseAccount.GetBlockedSenderCount():Integer;
begin
  Result := FBlockedSenders.Count;
end;

function TBaseAccount.GetSafeRecipientCount():Integer;
begin
  Result := FSafeRecipients.Count;
end;

destructor TBaseAccount.Destroy();
begin
  FreeAndNil(reg);
  FreeAndNil(FSafeSenders);
  FreeAndNil(FSafeRecipients);
  FreeAndNil(FBlockedSenders);
end;

function TBaseAccount.isStore(): Boolean;
begin
//  Result := (reg.ValueExists('MAPI Provider') and (reg.ReadInteger('MAPI Provider')=4));
  Result := length(DataFilePath) > 0;
end;

function TBaseAccount.IsAddressBook(): Boolean;
begin
  Result := (reg.ValueExists('MAPI Provider') and (reg.ReadInteger('MAPI Provider')=2));
end;

function TBaseAccount.IsMailbox(): Boolean;
begin
  Result := (reg.ValueExists('Email')) or ((reg.ValueExists('MAPI Provider') and (reg.ReadInteger('MAPI Provider')=5)));
end;

function TEmailSettings.FillInDetails(): TBaseAccount;
var
  detail : TBaseAccount;
  i : Integer;
  reg : TRegistry;
begin
  detail := nil;
  for i := 0 to outlooksubKeys2.Count - 1 do
  begin
    reg := TRegistry.Create;
    detail := TBaseAccount.Create('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\' + outlooksubKeys2.Strings[i]);
    emailDetails.Add(detail);
    try

    finally
      if Assigned(reg) then FreeAndNil(reg);
    end;
  end;

  for i :=0 to outlooksubKeys.Count-1 do
  begin
    reg := TRegistry.Create;
    detail := TBaseAccount.Create('\Software\Microsoft\Office\Outlook\OMI Account Manager\Accounts\' + outlooksubKeys.Strings[i]);
    emailDetails.Add(detail);
    try
      detail.AccessProgram := 'Outlook';
      if reg.ValueExists('SMTP Display Name') then
        detail.DisplayName := reg.ReadString('SMTP Display Name');
      if reg.ValueExists('SMTP Port') then
        detail.SMTPPort    := reg.ReadInteger('SMTP Port')
      else detail.SMTPPort := 25;
      if reg.ValueExists('SMTP Server') then
        detail.SMTPMailServer := reg.ReadString('SMTP Server');
      if reg.ValueExists('SMTP Email Address') then
        detail.EmailAddress := reg.ReadString('SMTP Email Address');
      if reg.ValueExists('POP3 Server') then
        detail.POP3Server       := reg.ReadString('POP3 Server');
      if reg.ValueExists('POP3 Port') then
        detail.POP3Port         := reg.ReadInteger('POP3 Port')
      else detail.POP3Port := 110;
    finally
      if Assigned(reg) then FreeAndNil(reg);
    end;
    detail := nil;
  end;

  for i := 0 to expresssubKeys.Count - 1 do
  begin
    try
      detail := TBaseAccount.Create('\Software\Microsoft\Internet Account Manager\Accounts\' + expresssubKeys.Strings[i]);
      detail.DisplayName      :='';
      detail.SMTPMailServer   :='';
      detail.SMTPPort         :=0;
      detail.EmailAddress := '';
      detail.POP3Server       := '';
      detail.POP3Port         := 0;
      detail.RegistryKey      := '';
      detail.AccessProgram := 'Outlook Express';
      emailDetails.Add(detail);
      detail.RegistryKey := '\Software\Microsoft\Internet Account Manager\Accounts\' + expresssubKeys.Strings[i];
      if reg.ValueExists('SMTP Display Name') then
        detail.DisplayName := reg.ReadString('SMTP Display Name');
      if reg.ValueExists('SMTP Port') then
        detail.SMTPPort    := reg.ReadInteger('SMTP Port')
      else detail.SMTPPort := 25;
      if reg.ValueExists('SMTP Server') then
        detail.SMTPMailServer := reg.ReadString('SMTP Server');
      if reg.ValueExists('SMTP Email Address') then
        detail.EmailAddress := reg.ReadString('SMTP Email Address');
      if reg.ValueExists('POP3 Server') then
        detail.POP3Server       := reg.ReadString('POP3 Server');
      if reg.ValueExists('POP3 Port') then
        detail.POP3Port         := reg.ReadInteger('POP3 Port')
      else detail.POP3Port := 110;
    finally
      if Assigned(reg) then FreeAndNil(reg);
    end;
    detail := nil;
  end;
end;

function TEmailSettings.FindEmailAccount(addr:String):TBaseAccount;
var
  i : Integer;
  detail : TBaseAccount;
begin
  Result := nil;
  for i := 0 to emailDetails.Count - 1 do
  begin
    detail := EmailDetails.Items[i];
    if CompareText(detail.EmailAddress, addr) = 0 then
    begin
      Result := EmailDetails.Items[i];
      break;
    end;
  end;
end;

{procedure TEmailSettings.Debug();
var
  i : Integer;
  detail : PEmailDetails;
begin
  for i:=0 to emailDetails.Count-1 do
    begin
      detail := emailDetails.Items[i];
      Writeln('---------------------------------------');
      Writeln('Email Address: '+detail.SMTPEmailAddress);
      Writeln('Display Name: '+ detail.DisplayName);
      Writeln('SMTP Server: ' + detail.SMTPMailServer +':'+
         IntToStr(detail.SMTPPort));
      Writeln('POP3 Server: ' + detail.POP3Server+':'+
         IntToStr(detail.POP3Port));
      Writeln('Registry Key: ' + detail.RegistryKey);
    end;
end;}

function TEmailSettings.Debug(): String;
var
  i : Integer;
  detail : TBaseAccount;
begin
  Result := '';
  for i := 0 to emailDetails.Count - 1 do
  begin
    detail := emailDetails.Items[i];
    Result := Result + '---------------------------------------' + crlf;
    Result := Result + 'Email Address: '+detail.EmailAddress  + crlf;
    Result := Result + 'Display Name: '+ detail.DisplayName  + crlf;
    Result := Result + 'SMTP Server: ' + detail.SMTPMailServer +':' + IntToStr(detail.SMTPPort) + crlf;
    Result := Result + 'POP3 Server: ' + detail.POP3Server+':' + IntToStr(detail.POP3Port) + crlf;
    Result := Result + 'Service UID: ' + detail.ServiceUID + crlf;
    Result := Result + 'Data File Path: ' + detail.DataFilePath + crlf;
    Result := Result + 'Registry Key: ' + detail.RegistryKey + crlf;
    Result := Result + 'isStore: ' + BoolToStr(detail.isStore) + crlf;
    Result := Result + 'isAddressBook: ' + BoolToStr(detail.IsAddressBook) + crlf;
    Result := Result + 'isMailbox: ' + BoolToStr(detail.IsMailbox) + crlf;
  end;
{$ifdef CONSOLE}
  Writeln(Result);
{$endif}
end;


constructor TEmailSettings.Create;
begin
  outlookAccounts     := TRegistry.Create;
  outlookAccounts2    := TRegistry.Create;
  expressAccounts     := TRegistry.Create;
  outlooksubKeys      := TStringList.Create;
  outlooksubKeys2     := TStringList.Create;
  expresssubKeys      := TStringList.Create;
  emailDetails        := TList.Create;

  outlookAccounts.RootKey  := HKEY_CURRENT_USER;
  outlookAccounts2.RootKey := HKEY_CURRENT_USER;
  expressAccounts.RootKey  := HKEY_CURRENT_USER;
  expressAccounts.OpenKey('Software\Microsoft\Internet Account Manager\Accounts', False);
// if(
  outlookAccounts.OpenKey('Software\Microsoft\Office\Outlook\OMI Account Manager\Accounts', False);//) then
  outlookAccounts2.OpenKey('Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676', False);//) then
//  begin
//    OutputDebugString('Registry Key opened');
//  end;
  outlookAccounts.GetKeyNames(outlooksubKeys);
  outlookAccounts2.GetKeyNames(outlooksubKeys2);
  expressAccounts.GetKeyNames(expresssubKeys);
  FillInDetails;
end;

destructor TEmailSettings.Destroy;
var
  i : Integer;
begin
  if Assigned(outlookAccounts) then FreeAndNil(outlookAccounts);
  if Assigned(outlookAccounts2) then FreeAndNil(outlookAccounts2);
  if Assigned(expressAccounts) then FreeAndNil(expressAccounts);
//  if Assigned(emailDetails) then FreeAndNil(emailDetails);
  if Assigned(outlooksubKeys) then FreeAndNil(outlooksubKeys);
  if Assigned(outlooksubKeys2) then FreeAndNil(outlooksubKeys2);
  if Assigned(expresssubKeys) then FreeAndNil(expresssubKeys);
  for i:=0 to emailDetails.Count-1 do
  begin
//      emailDetails.
//      FreeAndNil(emailDetails.Items[i]);
  end;
  if Assigned(emaildetails) then FreeAndNil(emaildetails);
end;

function TEmailSettings.GetEmailCount(): Integer;
var
  i : Integer;
begin
  OutputDebugString(PChar('subKeys Count : ' + IntToStr(expresssubKeys.Count)));
  for i := 0 to outlooksubKeys.Count-1 do
  begin
    OutputDebugString(PChar(IntToStr(i) + ' ' + outlooksubKeys.Strings[i]));
  end;
  for i := 0 to expresssubKeys.Count-1 do
  begin
    OutputDebugString(PChar(IntToStr(i) + ' ' + expresssubKeys.Strings[i]));
  end;
  Result := emailDetails.Count;
end;

end.

