unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  LYTray, Menus, StdCtrls, Buttons, ADODB,
  ActnList, AppEvnts, ComCtrls, ToolWin, ExtCtrls,
  registry,inifiles,Dialogs,
  StrUtils, DB, ComObj,Variants,ShellAPI;

type
  TfrmMain = class(TForm)
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    ApplicationEvents1: TApplicationEvents;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton8: TToolButton;
    ActionList1: TActionList;
    editpass: TAction;
    about: TAction;
    stop: TAction;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ToolButton7: TToolButton;
    SaveDialog1: TSaveDialog;
    ADOConn_BS: TADOConnection;
    Timer1: TTimer;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ApplicationEvents1Activate(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    procedure WMSyscommand(var message:TWMMouse);message WM_SYSCOMMAND;
    procedure UpdateConfig;{配置文件生效}
    function LoadInputPassDll:boolean;
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;EOT=#$4;ETB=#$17;
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  ifRecLog:boolean;//是否记录调试日志
  EquipUnid:integer;//设备唯一编号

  DaanConnStr:string;
  ifConnSucc:boolean;

  RFM:STRING;       //返回数据
  hnd:integer;
  bRegister:boolean;

{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ctext        :string;
  reg          :tregistry;
begin
  rfm:='';
  
  ConnectString:=GetConnectString;
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  Caption:='数据接收服务'+ExtractFileName(Application.ExeName);
  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);

//=============================初始化密码=====================================//
    reg:=tregistry.Create;
    reg.RootKey:=HKEY_CURRENT_USER;
    reg.OpenKey('\sunyear',true);
    ctext:=reg.ReadString('pass');
    if ctext='' then
    begin
        reg:=tregistry.Create;
        reg.RootKey:=HKEY_CURRENT_USER;
        reg.OpenKey('\sunyear',true);
        reg.WriteString('pass','JIHONM{');
        //MessageBox(application.Handle,pchar('感谢您使用智能监控系统，'+chr(13)+'请记住初始化密码：'+'lc'),
        //            '系统提示',MB_OK+MB_ICONinformation);     //WARNING
    end;
    reg.CloseKey;
    reg.Free;
//============================================================================//
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    if LoadInputPassDll then action:=cafree else action:=caNone;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    if not LoadInputPassDll then exit;
    application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  show;
end;

procedure TfrmMain.ApplicationEvents1Activate(Sender: TObject);
begin
  hide;
end;

procedure TfrmMain.WMSyscommand(var message: TWMMouse);
begin
  inherited;
  if message.Keys=SC_MINIMIZE then hide;
  message.Result:=-1;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  autorun:boolean;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);

  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');
  EquipUnid:=ini.ReadInteger(IniSection,'设备唯一编号',-1);

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  DaanConnStr:=ini.ReadString(IniSection,'连接达安数据库','');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  try
    ADOConn_BS.Connected := false;
    ADOConn_BS.ConnectionString := DaanConnStr;
    ADOConn_BS.Connected := true;
    ifConnSucc:=true;
  except
    on E:Exception do
    begin
      ifConnSucc:=false;
      MESSAGEDLG('连接达安数据库失败!'+E.Message,mtError,[mbOK],0);
    end;
  end;
end;

function TfrmMain.LoadInputPassDll: boolean;
TYPE
    TDLLFUNC=FUNCTION:boolean;
VAR
    HLIB:THANDLE;
    DLLFUNC:TDLLFUNC;
    PassFlag:boolean;
begin
    result:=false;
    HLIB:=LOADLIBRARY('OnOffLogin.dll');
    IF HLIB=0 THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    DLLFUNC:=TDLLFUNC(GETPROCADDRESS(HLIB,'showfrmonofflogin'));
    IF @DLLFUNC=NIL THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    PassFlag:=DLLFUNC;
    FREELIBRARY(HLIB);
    result:=passflag;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='连接达安数据库'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '设备唯一编号'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
  end;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  Memo1.Lines.Clear;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  SaveDialog1.DefaultExt := '.txt';
  SaveDialog1.Filter := 'txt (*.txt)|*.txt';
  if not SaveDialog1.Execute then exit;
  memo1.Lines.SaveToFile(SaveDialog1.FileName);
  showmessage('保存成功!');
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
VAR
  adotemp22,adotemp33,adotemp44,adotemp55:tadoquery;
  ReceiveItemInfo:OleVariant;
  FInts:OleVariant;
  sSex,sRemark,sAgeUnit:String;
  i,k:integer;
  Pathology_Type:String;//病理检测类型
begin
  if not ifConnSucc then exit;

  if length(memo1.Lines.Text)>=60000 then memo1.Lines.Clear;//memo只能接受64K个字符

  adotemp22:=tadoquery.Create(nil);
  adotemp22.Connection:=ADOConn_BS;
  adotemp22.Close;
  adotemp22.SQL.Clear;
  adotemp22.SQL.Text:='select * from da_outspecimen where status in (''4'',''5'') and createdate>GETDATE()-90';
  adotemp22.Open;
  while not adotemp22.Eof do
  begin
    memo1.Lines.Add('获取病人信息,requestcode:'+adotemp22.fieldbyname('requestcode').AsString+',patientname:'+adotemp22.fieldbyname('patientname').AsString);

    adotemp33:=tadoquery.Create(nil);
    adotemp33.Connection:=ADOConn_BS;
    adotemp33.Close;
    adotemp33.SQL.Clear;
    adotemp33.SQL.Text:='select datestcode as item_code,testresult as item_result,reportremark as remark from da_result where requestcode='''+adotemp22.fieldbyname('requestcode').AsString+''' and status=''1'' and isnull(datestcode,'''')<>'''' '+//普通检验项目结果
                        ' union all '+
                        'select anticode as item_code,isnull(resultvalue,'''')+''   ''+isnull(testresult,'''') as item_result,'''' as remark from da_micantiresult where requestcode='''+adotemp22.fieldbyname('requestcode').AsString+''' '+
                        ' union all '+
                        'select organismcode as item_code,quantity as item_result,quantitycomment as remark from da_micorgresult where requestcode='''+adotemp22.fieldbyname('requestcode').AsString+''' and status=''2'' '+
                        ' union all '+
                        'select itemname as item_code,result as item_result,'''' as remark from da_pathologyresult where requestcode='''+adotemp22.fieldbyname('requestcode').AsString+''' and isnull(result,'''')<>'''' ';//病理检测结果
    adotemp33.Open;
    
    if adotemp33.RecordCount<=0 then begin adotemp33.Free;adotemp22.Next;continue;end;
    
    adotemp55:=tadoquery.Create(nil);
    adotemp55.Connection:=ADOConn_BS;
    adotemp55.Close;
    adotemp55.SQL.Clear;
    adotemp55.SQL.Text:='select top 1 * from da_outspecimen do,da_pathologyresult dp where do.requestcode=dp.requestcode and do.requestcode='''+adotemp22.fieldbyname('requestcode').AsString+''' and isnull(dp.result,'''')<>'''' ';
    adotemp55.Open;
    if adotemp55.RecordCount>0 then//存在病理检测结果
    begin
      k:=1;
      Pathology_Type:=adotemp55.fieldbyname('datestnames').AsString;//病理检测类型
    end;
    adotemp55.Free;

    sRemark:=adotemp22.fieldbyname('remark').AsString;    

    ReceiveItemInfo:=VarArrayCreate([0,adotemp33.RecordCount-1+k],varVariant);

    i:=0;
    while not adotemp33.Eof do
    begin    
      memo1.Lines.Add('获取病人结果,item_code:'+adotemp33.fieldbyname('item_code').AsString+',item_result:'+adotemp33.fieldbyname('item_result').AsString);

      sRemark:=sRemark+adotemp33.fieldbyname('remark').AsString;

      ReceiveItemInfo[i]:=VarArrayof([adotemp33.fieldbyname('item_code').AsString,adotemp33.fieldbyname('item_result').AsString,'','']);

      inc(i);
      adotemp33.Next;
    end;
    adotemp33.Free;

    if k>0 then ReceiveItemInfo[i]:=VarArrayof(['病理检测类型',Pathology_Type,'','']);

    if adotemp22.fieldbyname('sex').AsString='M' THEN sSex:='男'
      else if adotemp22.fieldbyname('sex').AsString='F' THEN sSex:='女'
        else sSex:='未知';

    if adotemp22.fieldbyname('ageunit').AsString='1' THEN sAgeUnit:='月'
      else if adotemp22.fieldbyname('ageunit').AsString='2' THEN sAgeUnit:='天'
        else if adotemp22.fieldbyname('ageunit').AsString='3' THEN sAgeUnit:='小时'
          else sAgeUnit:='岁';

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,adotemp22.fieldbyname('requestcode').AsString,
        FormatDateTime('YYYY-MM-DD hh:nn:ss',adotemp22.fieldbyname('enterbydate').AsDateTime),
        (GroupName),(SpecType),adotemp22.fieldbyname('samstate').AsString,(EquipChar),
        (CombinID),
        adotemp22.fieldbyname('patientname').AsString+'{!@#}'+sSex+'{!@#}{!@#}'+adotemp22.fieldbyname('age').AsString+sAgeUnit+'{!@#}'+adotemp22.fieldbyname('patientnumber').AsString+'{!@#}'+adotemp22.fieldbyname('location').AsString+'{!@#}'+adotemp22.fieldbyname('doctor').AsString+'{!@#}'+adotemp22.fieldbyname('bednumber').AsString+'{!@#}'+adotemp22.fieldbyname('diagnostication').AsString+'{!@#}'+copy(sRemark,1,50)+'{!@#}{!@#}{!@#}'+adotemp22.fieldbyname('hospsampleid').AsString+'{!@#}',
        (LisFormCaption),(ConnectString),
        (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
        ifRecLog,true,'常规',
        '',
        EquipUnid,
        '','','','',
        -1,-1,-1,-1,
        -1,-1,-1,-1,
        false,false,false,false);
      if not VarIsEmpty(FInts) then FInts:= unAssigned;
    end;
    
    adotemp44:=tadoquery.Create(nil);
    adotemp44.Connection:=ADOConn_BS;
    adotemp44.Close;
    adotemp44.SQL.Clear;
    adotemp44.SQL.Text:='update da_outspecimen set status=''9'' where requestcode='''+adotemp22.fieldbyname('requestcode').AsString+''' ';
    adotemp44.ExecSQL;
    adotemp44.Free;

    adotemp22.Next;
  end;
  adotemp22.Free;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.




        

