[Setup]
AppName=Aride
AppVerName=Aride
DefaultDirName={pf}\Aride
DefaultGroupName=Aride

[Languages]
Name: french; MessagesFile: compiler:Languages\French.isl

[Files]
; [Bootstrap Files]
; @COMCAT.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,6/1/98 12:00:00 AM,22288,4.71.1460.1
Source: DLL\COMCAT.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @MSVCRT40.DLL,$(WinSysPathSysFile),,,6/1/98 12:00:00 AM,326656,4.21.0.0
Source: DLL\MSVCRT40.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @VB6FR.DLL,$(WinSysPath),,$(Shared),10/4/06 5:08:38 PM,119568,6.0.89.88
Source: DLL\VB6FR.DLL; DestDir: {sys}; Flags: promptifolder sharedfile
; @stdole2.tlb,$(WinSysPathSysFile),$(TLBRegister),,7/14/09 12:43:53 AM,16896,6.1.7600.16385
;Source: DLL\stdole2.tlb; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
; @asycfilt.dll,$(WinSysPathSysFile),,,3/5/10 8:42:42 AM,67584,6.1.7600.16544
Source: DLL\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @olepro32.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,7/14/09 2:16:12 AM,90112,6.1.7600.16385
Source: DLL\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @oleaut32.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,8/27/11 5:43:07 AM,571904,6.1.7600.16872
Source: DLL\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @msvbvm60.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,7/14/09 2:15:50 AM,1386496,6.0.98.15
Source: DLL\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver

; [Setup1 Files]
; @MSCMCFR.DLL,$(WinSysPath),,$(Shared),7/13/98 12:00:00 AM,141312,6.0.81.63
; UNKNOWN
;;Source: DLL\MSCMCFR.DLL; DestDir: {sys}; Flags: promptifolder sharedfile
; @MSCOMCTL.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),6/6/12 7:59:42 PM,1070152,6.1.98.34
; Common controls display like scrollbar (useful for updater)
Source: DLL\MSCOMCTL.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @INETFR.DLL,$(WinSysPath),,$(Shared),7/13/98 12:00:00 AM,15360,6.0.81.63
;;;Source: DLL\INETFR.DLL; DestDir: {sys}; Flags: promptifolder sharedfile
; @MSINET.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),10/4/06 5:02:17 PM,132880,6.1.97.82
Source: DLL\MSINET.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @TABCTL3N.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),12/5/02 6:58:24 PM,209608,6.0.90.43
;#Source: DLL\TABCTL3N.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @RCHTXFR.DLL,$(WinSysPath),,$(Shared),7/13/98 12:00:00 AM,34304,6.0.81.63
;;;Source: DLL\RCHTXFR.DLL; DestDir: {sys}; Flags: promptifolder sharedfile
; @RICHED32.DLL,$(WinSysPathSysFile),,,5/8/98 12:00:00 AM,174352,4.0.993.4
;Source: DLL\RICHED32.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @RICHTX32.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),5/22/00 5:00:00 AM,203976,6.0.88.4
;#Source: DLL\RICHTX32.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @MSWINSCN.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),12/5/02 6:58:04 PM,109248,6.0.88.4
; UNKNOWN but useful
Source: DLL\MSWINSCN.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @scrrnfr.dll,$(WinSysPath),,$(Shared),8/5/04 1:00:00 PM,24626,5.6.0.6626
; UNKNOWN
;;Source: DLL\scrrnfr.dll; DestDir: {sys}; Flags: promptifolder sharedfile
; @msvcrt.dll,$(WinSysPathSysFile),,,12/16/11 8:59:17 AM,690688,7.0.7600.16930
; Maybe for OLE use
;;Source: DLL\msvcrt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @scrrun.dll,$(WinSysPath),$(DLLSelfRegister),$(Shared),7/14/09 2:16:13 AM,163840,5.8.7600.16385
; Some scripting object like dictionnary
;;;Source: DLL\scrrun.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @PaintX.dll,$(WinSysPath),$(DLLSelfRegister),$(Shared),3/7/02 12:19:16 AM,454656,1.0.5.0
; Print color on pixel
;;;Source: DLL\PaintX.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @wmp.dll,$(AppPath),,,9/1/10 5:29:28 AM,11406848,12.0.7600.16667
; Windows Media Player
;;;Source: DLL\wmp.dll; DestDir: {app}; Flags: promptifolder
; @dx7vb.dll,$(WinSysPath),$(DLLSelfRegister),$(Shared),8/5/04 1:00:00 PM,619008,5.3.2600.2180
; DirectX 7
;;;Source: DLL\dx7vb.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @vb5db.dll,$(WinSysPath),,$(Shared),6/17/98 11:00:00 PM,89360,6.0.81.69
;Source: DLL\vb5db.dll; DestDir: {sys}; Flags: promptifolder sharedfile
; @msrepl35.dll,$(WinSysPathSysFile),,,8/25/99 2:57:26 PM,415504,3.51.3225.0
;;Source: DLL\msrepl35.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @msrd2x35.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,6/1/98 2:37:00 PM,262144,3.51.623.0
;;Source: DLL\msrd2x35.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @expsrv.dll,$(WinSysPathSysFile),,,7/14/09 2:15:20 AM,380957,6.0.72.9589
;;Source: DLL\expsrv.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @vbajet32.dll,$(WinSysPathSysFile),,,7/14/09 2:16:17 AM,30749,6.0.1.9431
;;Source: DLL\vbajet32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @MSJINT35.DLL,$(WinSysPathSysFile),,,7/7/98 12:00:00 AM,149776,3.51.623.0
;;Source: DLL\MSJINT35.DLL; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @msjter35.dll,$(WinSysPathSysFile),,,6/10/99 9:34:04 AM,24848,3.51.623.0
;;Source: DLL\msjter35.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @msjet35.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,9/28/99 9:42:48 PM,1050896,3.51.3328.0
;;Source: DLL\msjet35.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @DAO350.DLL,$(MSDAOPath),$(DLLSelfRegister),$(Shared),4/26/98 11:00:00 PM,570128,3.51.1608.0
;Source: DLL\DAO350.DLL; DestDir: {dao}; Flags: promptifolder regserver sharedfile
; @aamd532.dll,$(WinSysPath),,$(Shared),4/17/99 11:06:40 AM,10752,1.0.0.1
; Use for MD5 calc
Source: DLL\aamd532.dll; DestDir: {sys}; Flags: promptifolder sharedfile
; @IPHLPAPI.DLL,$(WinSysPath),,$(Shared),7/14/09 2:15:33 AM,103936,6.1.7600.16385
; Use for TCP exchange
;;;Source: DLL\IPHLPAPI.DLL; DestDir: {sys}; Flags: promptifolder sharedfile
; @ws2_32.dll,$(WinSysPath),,$(Shared),7/14/09 2:16:20 AM,206336,6.1.7600.16385
; Winsock V2
;;;Source: DLL\ws2_32.dll; DestDir: {sys}; Flags: promptifolder sharedfile
; @wininet.dll,$(WinSysPath),,$(Shared),11/14/12 2:57:37 AM,1129472,9.0.8112.16457
;Source: DLL\wininet.dll; DestDir: {sys}; Flags: promptifolder sharedfile
; @urlmon.dll,$(WinSysPath),$(DLLSelfRegister),$(Shared),11/14/12 2:57:44 AM,1103872,9.0.8112.16457
;Source: DLL\urlmon.dll; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @Client.exe,$(AppPath),,,1/12/13 6:03:27 PM,11452416,0.0.0.1
Source: EXE\Client.exe; DestDir: {app}; Flags: promptifolder

[Icons]
Name: {group}\Aride; Filename: {app}\Client.exe; WorkingDir: {app}

[Dirs]
Name: {app}\; Permissions: everyone-full

[Run]
Filename: {app}\Client.exe; WorkingDir: {app}; Flags: postinstall runascurrentuser; Description: Lancer Aride Online

[Code]
procedure Explode(var Dest: TArrayOfString; Text: String; Separator: String);
var
	i: Integer;
begin
	i := 0;
	repeat
		SetArrayLength(Dest, i+1);
		if Pos(Separator,Text) > 0 then	begin
			Dest[i] := Copy(Text, 1, Pos(Separator, Text)-1);
			Text := Copy(Text, Pos(Separator,Text) + Length(Separator), Length(Text));
			i := i + 1;
		end else begin
			 Dest[i] := Text;
			 Text := '';
		end;
	until Length(Text)=0;
end;

procedure OppositeRmTree(Directory: String);
begin
	while True do
		if DirExists(Directory) then
		begin
			if Directory = ExpandConstant('{app}') then
			begin
				Break;
			end
			else
			begin
				if RemoveDir(Directory) then
				begin
					Directory := ExtractFileDir(Directory);
				end
				else
					Break;
			end;
		end
		else
		begin
			Break;
		end;
end;

procedure RmVersion(VersionFile: String);
var
	i: Integer;
	arrayLen: Longint;
	item: String;
	oldItemDirectory: String;
	itemsToDelete: Array of String;
	List: TArrayOfString;
begin
  if FileExists(VersionFile) then
  begin
    LoadStringsFromFile(VersionFile, itemsToDelete);
    arrayLen := GetArrayLength(itemsToDelete);
    oldItemDirectory := '';
    for i := arrayLen-1 downto 0 do
    begin
      //List := TArrayOfString.Create;
      //try
      Explode(List, itemsToDelete[i], '=');

      if i <> arrayLen-1 then
      begin
        if ExtractFileDir(ExpandConstant('{app}') + List[0]) <> oldItemDirectory then
        begin
          OppositeRmTree(oldItemDirectory);
        end;
      end;
      oldItemDirectory := ExtractFileDir(ExpandConstant('{app}') + List[0]);

      if i = 0 then
        OppositeRmTree(oldItemDirectory);

      //ExtractStrings(['='], [], PChar(itemsToDelete[i]), List);
      item := ExpandConstant('{app}') + List[0];
      if FileExists(item) then
        DeleteFile(item);
      if DirExists(item) then
        RemoveDir(item);
      item := item + '.new'
      if FileExists(item) then
        DeleteFile(item);
      if DirExists(item) then
        RemoveDir(item);
    end;

    DeleteFile(VersionFile);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
	versionFile: String;
begin
  case CurUninstallStep of
    usUninstall:
      begin
		versionFile := ExpandConstant('{app}') + '\Version.ini';
		RmVersion(versionFile);
		RmVersion(versionFile + '.new');
	  end;
  end;
end;
