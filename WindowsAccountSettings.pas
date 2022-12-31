unit WindowsAccountSettings;

interface

uses
    windows
  , Registry
  , System.Generics.Collections
  , System.SysUtils
  , System.Classes
  ;

type
  TWindowsUserProfile = class
    private
      reg : TRegistry;
      function GetProfilePath(): String;
    public
      constructor Create(key: String);
      destructor Destroy; override;
      property ProfilePath: String read GetProfilePath;
  end;

  TWindowsProfiles = class
    private
      reg: TRegistry;
      FAccounts: TObjectList<TWindowsUserProfile>;
      function GetAccountByIndex(i: Integer): TWindowsUserProfile;
    public
      constructor Create;
      destructor Destroy; override;
      function Count(): Integer;
      property Accounts[i: Integer]: TWindowsUserProfile read GetAccountByIndex; default;
  end;

  procedure LoadUserHive(sPathToUserHive: string);

implementation


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

function TWindowsProfiles.GetAccountByIndex(i: Integer): TWindowsUserProfile;
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

end.
