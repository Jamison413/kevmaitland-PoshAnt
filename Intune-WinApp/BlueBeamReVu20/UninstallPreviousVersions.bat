TASKKILL /F /IM "BBPrint.exe"

REM If upgrading 10.0 or 10.1, and you installed to a custom path, please fix the INSTALLDIR path below:
SET INSTALLDIR=%ProgramFiles%\Bluebeam Software\Bluebeam Revu
SET ADMINPATH=Pushbutton PDF\PbMngr5.exe
SET VUADMINPATH=Bluebeam Vu Admin.exe

SET WCV=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\InstallShield
SET WCV64=HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\InstallShield
SET WCVF=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders
SET WGF=%WINDIR%\Installer
SET ISII=%ProgramFiles%\InstallShield Installation Information
SET ISII64=%ProgramFiles(x86)%\InstallShield Installation Information

SET UNINSTALLKEY=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall

REM On 64-bit systems, eXtreme registers in the 32-bit node in the registry.
SET UNINSTALLKEY64=HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall

REM ==========================================
REM Bluebeam v20 Release 20.2.60 
REM ==========================================

REM Uninstall Bluebeam Revu 20.2.60: 
SET GUID={3A415AE3-546A-47B4-9458-79C8C5751116}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.2.60 x64:
SET GUID={8070889A-41E8-497E-98B0-558FE00FFD9C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.2.50 
REM ==========================================

REM Uninstall Bluebeam Revu 20.2.50: 
SET GUID={5C9CF738-1918-4E6B-9A89-BAAD1C5C789C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.2.50 x64:
SET GUID={633857BC-A874-4209-A6B4-D826B3F7747B}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.2.40 
REM ==========================================

REM Uninstall Bluebeam Revu 20.2.40: 
SET GUID={229E38D2-B275-4152-97F0-F77751768A2E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.2.40 x64:
SET GUID={74F5FA28-F71D-48AF-9CF8-1722F9FE39A1}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.2.30 
REM ==========================================

REM Uninstall Bluebeam Revu 20.2.30: 
SET GUID={6B49A166-DD4E-48FD-9630-F5632225B7EB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.2.30 x64:
SET GUID={88B4346C-8BF8-4144-BA5B-AD1504CAB5AE}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.2.20 
REM ==========================================

REM Uninstall Bluebeam Revu 20.2.20: 
SET GUID={17763849-D695-4F56-AD43-3083DE60FA3F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.2.20 x64:
SET GUID={4D4C2614-EFD5-49CD-A28A-A41D7C42C74B}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.2.15 
REM ==========================================

REM Uninstall Bluebeam Revu 20.2.15: 
SET GUID={31B1B1FA-A571-452A-B063-6CAF8A786725}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.2.15 x64:
SET GUID={8AC808CF-ADE1-4417-A737-89929725A6D2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.1.100
REM ==========================================

REM Uninstall Bluebeam Revu 20.1.100: 
SET GUID={4B7AB0B8-B8B5-4E87-B6A6-F2D8A593E5B2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.1.100 x64:
SET GUID={63B74013-ED33-4B6E-9B26-1704883A4BD2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.1.20 
REM ==========================================

REM Uninstall Bluebeam Revu 20.1.20: 
SET GUID={1C104CE3-1683-4045-AD2D-8A9EDF2844F1}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.1.20 x64:
SET GUID= {32ED3670-111B-41DF-9E3C-C63C0933A1A0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.1.15 
REM ==========================================

REM Uninstall Bluebeam Revu 20.1.15: 
SET GUID={2626206D-A317-48AF-B113-881E381EACD6}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.1.15 x64:
SET GUID={43117D4A-7701-4A14-936D-24C188436D3D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.0.30 
REM ==========================================

REM Uninstall Bluebeam Revu 20.0.30: 
SET GUID={27DB9266-B677-4391-8AEA-0F77DB037E23}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.0.30 x64:
SET GUID={633A858B-F891-45BC-A355-46AB41C7D310}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn


REM ==========================================
REM Bluebeam v20 Release 20.0.20 
REM ==========================================

REM Uninstall Bluebeam Revu 20.0.20: 
SET GUID={52BF0274-B323-48D6-80D4-8DEBAABADAA1}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.0.20 x64:
SET GUID={9CB75A38-6CD7-4968-A339-58211FB11C85}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam OCR 20 Release 20.0.2 
REM ==========================================

REM Uninstall Bluebeam OCR 20.0.2: 
SET GUID={6A72B9F2-5936-44E7-A3D0-A2E13A9F88BB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam OCR 20.0.2 x64:
SET GUID={3688D52C-2CB4-4216-8D15-4978D28A5BBF}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v20 Release 20.0.15 
REM ==========================================

REM Uninstall Bluebeam Revu 20.0.15: 
SET GUID={593224EF-3F1A-4FD7-B586-017147CE2FDD}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.0.15 x64:
SET GUID={1C567966-34EF-43F5-8C00-EF0493B3763E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam Revu 20 Release 20.0.2 Beta 2
REM ==========================================

REM Uninstall Bluebeam Revu 20.0.2: 
SET GUID={80A9B7AD-A601-44D0-A3B2-D24EA373930C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.0.2 x64:
SET GUID={79B174F0-C4E2-409C-BD37-9A70E3F02462}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam Revu 20 Release 20.0.1 Beta 1
REM ==========================================

REM Uninstall Bluebeam Revu 20.0.1: 
SET GUID={FD4649E5-1C38-4C28-B33A-8CE125007CDB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 20.0.1 x64:
SET GUID={A3E33A5E-92CA-4218-817B-83DBB95FCAFE}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam OCR 20 Release 20.0.1 
REM ==========================================

REM Uninstall Bluebeam OCR 20.0.1: 
SET GUID={1A0CA402-8537-4755-872C-B6CB8BA8490C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam OCR 20.0.1 x64:
SET GUID={B4E5CE58-FB32-4201-B1D6-AFEBE203ED38}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.2.15
REM ==========================================

REM Uninstall Bluebeam Revu 19.2.15: 
SET GUID={5ADA335F-9AEE-4FBC-B68C-E080547904AE}

REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.2.15 x64:
SET GUID={80E715A1-92E3-4D65-8E32-95D4A4290323}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.1.3 Beta 3
REM ==========================================

REM Uninstall Bluebeam Revu 19.1.3: 
SET GUID={41B010D6-7A6D-4FBC-A8A1-C38A55A37B65}

REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.1.3 x64:
SET GUID={79D6F3AA-8AE7-48EC-B20D-8441CA5B34B8}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19.1 Beta 2
REM ==========================================

REM Uninstall Bluebeam Revu 19.1.2: 
SET GUID={1038A5D5-70CC-4BA0-9374-DFBB31B13A63}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.1.2 x64:
SET GUID={5247251C-FF43-492D-9EEB-45F01D4F200D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19.1 Beta 1
REM ==========================================

REM Uninstall Bluebeam Revu 19.1.1: 
SET GUID={321891E1-AD99-4619-BDBC-545DD94A1DFA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.1.1 x64:
SET GUID={65149FB8-2513-4DE4-B497-B00FC6DDDA45}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Base Release 19.0.15 Master, 19.0.20 19.1.15, 19.1.16, 19.1.20
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.15: 
SET GUID={4BB92938-EBC4-45F1-A3AE-E1EB80574DA9}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.15 x64:
SET GUID={74A435D8-FE24-498A-9809-76541D74182A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.0.8 Beta 8
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.8: 
SET GUID={09361F62-ABC4-44CB-8631-0711CCC562B2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.8 x64:
SET GUID={29862695-123C-4096-9535-B3279A2D5B0E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.0.6 Beta 6 & 7
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.6: 
SET GUID={57E360B7-D787-43DF-A9CD-301D8C6B19F3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.6 x64:
SET GUID={98A417A6-4E8A-4FBA-9880-9883AB882627}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.0.5 Beta 5
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.5: 
SET GUID={4C1A9354-DDD1-4F04-BE9E-FCBA7C6DD852}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.5 x64:
SET GUID={59C37150-DACB-42B2-8AD7-ADC59F9FDBEA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.0.4 Beta 4
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.4: 
SET GUID={3F88B050-11D5-4BAD-8BE2-23B11725A1E9}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.4 x64:
SET GUID={60EF9BC0-9953-4B4B-B3E1-5D02E4BC4165}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.0.3 Beta 3
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.3: 
SET GUID={3081FBE4-1F79-4C84-9282-F673B960C3FB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.3 x64:
SET GUID={6F907076-8BC9-417B-8209-F91B0E96ADB0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 Release 19.0.2 Beta 2
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.2: 
SET GUID={1BC449F6-3B05-471C-82AC-6FDFF0004744}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.2 x64:
SET GUID={8352D4CE-50EF-45FA-BFD3-4A64BA3FC8D8}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v19 BRelease 19.0.1 Beta 1
REM ==========================================

REM Uninstall Bluebeam Revu 19.0.1: 
SET GUID={E9558AC7-D6BF-44C6-980B-EC176C661623}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 19.0.1 x64:
SET GUID={87177218-CF4A-4EF5-AC72-99BA42F66B4A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v18 Base Release 18.7.0 
REM ==========================================

REM Uninstall Bluebeam Revu 18.7.0: 
SET GUID={17F7A5C6-E209-4A58-B902-46AF8363C2E4}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 18.7.0 x64:
SET GUID={5682AD5E-66AD-4267-93CC-8D90CD78EF4E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v18 Base Release 18.2.1 (Japan)
REM ==========================================

REM Uninstall Bluebeam Revu 18.2.1: 
SET GUID={30E3091F-7AAF-42CD-88A8-227E698ABF06}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 18.2.1 x64:
SET GUID={459B4880-C6AB-4076-989E-6F190ABB2066}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v18 Base Release 18.0.3
REM ==========================================

REM Uninstall Bluebeam Revu 18.0.3: 
SET GUID={6F6332BC-7C57-4611-B433-8E90955BF6D5}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 18.0.3 x64:
SET GUID={7F5E49F6-A466-4553-B9E0-53D7380944E3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v18 Beta 2
REM ==========================================

REM Uninstall Bluebeam Revu 18.0.2: 
SET GUID={56465A31-194B-42A9-B0BB-30FE23A69EEA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 18.0.2 x64:
SET GUID={6C5C5113-BEA3-4D55-B55F-0FDD033AE159}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn


REM ==========================================
REM Bluebeam v18 Beta 1
REM ==========================================

REM Uninstall Bluebeam Revu 18.0.1: 
SET GUID={AF5E7F9C-5B10-42CE-8461-BB3D8376D80E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 18.0.1 x64:
SET GUID={B3AAB52B-0671-42A4-8024-132218BD759F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v18 
REM ==========================================

REM Uninstall Bluebeam Revu 18.0.0: 
SET GUID={E6F96FCB-B7DC-4773-8D4C-439DF2715FE2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 18.0.0 x64:
SET GUID={14059DC8-0F22-4FC3-BED7-1E23BD5C01EA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17.0.40
REM ==========================================

REM Uninstall Bluebeam Vu 17.0.40:
SET GUID={8E8256CA-6ED4-435D-9903-14F9A7DD4DC2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.40 x64:
SET GUID={AAAE59F3-C340-4AEF-A091-C2F83D650BD3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17.0.30
REM ==========================================

REM Uninstall Bluebeam Vu 17.0.30:
SET GUID={23889E29-FB73-4991-BDA2-AEA88554A6CD}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.30 x64:
SET GUID={5659D298-7DE7-48DD-A9E5-6551EDE2535E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17.0.20 Patch Update 2
REM ==========================================

REM Uninstall Bluebeam Vu 17.0.20:
SET GUID={B2581DBA-ED08-4846-831E-45F736A1D6D4}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.20 x64:
SET GUID={60D0BC0C-CC45-4CB4-B7AC-643AA08251C1}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17.0.10 Patch Update 1
REM ==========================================

REM Uninstall Bluebeam Vu 17.0.10:
SET GUID={1CEDDDAF-ED32-4513-A90A-DA07CAAF302F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.10 x64:
SET GUID={94A0E313-E73B-4DB0-8C2F-3F40A7155EFC}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17 Master
REM ==========================================

REM Uninstall Bluebeam Revu 17.0.8: 
SET GUID={4681519E-109B-4BD5-90E2-83DBB8720518}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 17.0.8 x64:
SET GUID={A9FF6312-66C3-4D99-AA3F-40611C2360FD}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 17.0.8:
SET GUID={820705B3-A349-4A70-B702-02CC28334BB1}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.8 x64:
SET GUID={2C355124-A14C-4814-92EB-EF7FE71E97C0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17 RC
REM ==========================================

REM Uninstall Bluebeam Revu 17.0.7 RC: 
SET GUID={7D64B10E-6B02-4DAA-9E24-6FB697035C13}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 17.0.7 RC x64:
SET GUID={42DCB289-FF0E-4129-B258-738052AF34E0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 17.0.7:
SET GUID={0B7842E8-BDCB-4759-A8DA-D722904B88D3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.7 x64:
SET GUID={F6C9F283-5508-427A-9872-C83755507E0E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17 Beta 5
REM ==========================================

REM Uninstall Bluebeam Vu 17.0.6:
SET GUID={7BFBB9DB-6BDD-48DB-97B6-F511C7328655}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 17.0.6 x64:
SET GUID={4BA0B455-CD71-4D06-A02F-5EC98DB90D81}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM ==========================================
REM Bluebeam v17 Beta 1-4
REM ==========================================

REM Uninstall Bluebeam Revu 17.0.x Beta: 
SET GUID={AF843750-5933-4ADD-9B3D-9BE628364569}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 17.0.x x64 Beta:
SET GUID={69D988F0-04AF-42C7-98E8-06720F6DC0BE}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 17.0.x Beta:
SET GUID={87F01787-83CD-4FC7-8F4F-6A79B7786400}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 17.0.x x64 Beta:
SET GUID={E02674C1-E217-4352-9024-9C21BDDB123F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================
REM Bluebeam v16.5.x Release
REM =========================================

REM Uninstall Bluebeam Revu 16.5.0: 
SET GUID={B9A7B7BC-4FDB-469F-A862-60638E50F3B0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.5.0 x64:
SET GUID={BF1AAA9A-6B8A-4280-B7EE-C1116FB19B85}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.5.0 Localization package:  
SET GUID={C1E3C28A-E89E-4821-B103-28A6C5B472F9}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.5.0 x64 Localization package:
SET GUID={50788C61-F3F5-4199-B04B-3DCD8B314616}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.5.0:
SET GUID={CA142145-917F-4362-BC4B-919DEC6D3F38}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.5.0 x64:
SET GUID={01F30714-6AFE-4D5F-973B-C33673F8636D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.5.0 Localization package:  
SET GUID={A7B58689-A6B9-400B-8B45-C6122A8FD2EB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.5.0 x64 Localization package:
SET GUID={A7ED7959-DEB9-49E4-A597-13083A5BEC1C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================
REM Bluebeam v16.1 Release
REM =========================================

REM Uninstall Bluebeam Revu 16.1.0: 
SET GUID={E1BA1B26-8488-4C91-A6F8-399C9FCC2CF0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.1.0 x64:
SET GUID={50464486-13F5-41CA-AF25-AD56C0DC1D02}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.1.0 Localization package:  
SET GUID={C21FF26F-71A8-4C4C-BF14-5FC8B2EBFFC5}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.1.0 x64 Localization package:
SET GUID={2626F549-DAE5-4838-BB4E-347C4B81487F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.1.0:
SET GUID={9C0623E4-A5E1-4690-90FB-135210965826}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.1.0 x64:
SET GUID={F099023B-8463-4CD8-9815-04BA5C2FD2C3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.1.0 Localization package:  
SET GUID={79E88AA9-6F72-4E50-9C31-557BDF0D36D3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.1.0 x64 Localization package:
SET GUID={F09C3F26-865F-468A-BC0A-9667B038A080}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================
REM Bluebeam v16 Master Release
REM =========================================

REM Uninstall Bluebeam Revu 16.0.4: 
SET GUID={21244472-3597-484D-BACB-7D093E97630D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 16.0.4 x64:
SET GUID={B7D0D8F8-CCF4-4199-9593-351FC71C8483}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.4 Localization package:  
SET GUID={63405EFA-D251-486E-B4EF-A9B62FF312D6}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.4 x64 Localization package:
SET GUID={F17DC148-CCF0-4734-8C5E-3D57A14DDE15}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.4:
SET GUID={D337161E-8B1B-48A0-BC1A-5E33CA831117}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.4 x64:
SET GUID={0801B6A0-3DAD-4C30-B61F-3C17FB866095}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.4 Localization package:  
SET GUID={5FDA9A4F-A100-4C18-9A40-83E5CAF801D5}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.4 x64 Localization package:
SET GUID={84B6724E-2E4D-45AB-85A1-FCCE8FFAB26A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn


REM =========================================
REM Bluebeam v16 RC
REM =========================================

REM Uninstall Bluebeam Revu 16.0.3: 
SET GUID={20EC2EFC-EFBF-4E8D-8BA2-E3DEBB82B04C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 16.0.3 x64:
SET GUID={0691458F-48D6-4AC6-8DBC-299DE7BE40D7}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.3 Localization package:  
SET GUID={68C0AAAE-C746-4229-860C-B24AF5E5B8B4}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.3 x64 Localization package:
SET GUID={58BB16D9-FD8D-4595-A90A-19778CBF4A42}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.3:
SET GUID={77FCC43E-1E80-4DAA-BC80-C8D253563213}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.3 x64:
SET GUID={0B2AA47D-60D9-4283-962A-36903DCA9599}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.3 Localization package:  
SET GUID={744F0668-B452-481A-853F-FA433FDF25F2}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.3 x64 Localization package:
SET GUID={5EFBA230-8324-46E5-9DEC-BD0F23A2FF49}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn


REM =========================================
REM Bluebeam v16 beta 2
REM =========================================

REM Uninstall Bluebeam Revu 16.0.2: 
SET GUID={A4C3079A-53EB-4621-A484-6F6D449E91BF}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.2 x64:
SET GUID={DA2EF300-582D-4904-B3B5-965F8A7DCD22}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.2 Localization package:  
SET GUID={2C56258B-D518-4705-BBC9-B266AFD8F49C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.2 x64 Localization package:
SET GUID={3D4B8DCA-F4A7-4C03-8B7E-664D6657674D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.2:
SET GUID={1639F525-B17F-42F8-8EBF-9D47D8D346EA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.2 x64:
SET GUID={EBF6D777-1491-4596-ADA4-9567CC1B9966}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.2 Localization package:  
SET GUID={3386E4F7-7FA8-42A2-8133-2A6E998686AA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.2 x64 Localization package:
SET GUID={818B4230-687D-447C-BB4C-4C0B42CECDAB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM===========================================
REM Bluebeam v16 beta 1
REM =========================================

REM Uninstall Bluebeam Revu 16.0.1: 
SET GUID={AC90D392-BCBE-462A-BF9B-0E5D1F516F2F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.1 x64:
SET GUID={C3BEB75D-6D48-4FDC-AFA3-62CCF060BF2C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.1 Localization package:  
SET GUID={86794778-7E31-4155-9CA1-07FEB294489C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 16.0.1 x64 Localization package:
SET GUID={565AA50A-B49E-4AAC-99E9-7A1A66521258}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.1:
SET GUID={977A69DD-A2DA-4996-8117-B15C302B7D3B}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.1 x64:
SET GUID={0E97396E-1293-484C-8EA8-730F15260FAA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.1 Localization package:  
SET GUID={7186FB8E-95BB-4F30-BC4A-7250B9CF0ADC}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 16.0.1 x64 Localization package:
SET GUID={DD9E2900-B962-4B95-8F57-B4D68E82D8EC}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================
REM Bluebeam 2015 v15
REM =========================================

REM Uninstall Bluebeam Revu 15.6.0: 
SET GUID={D933E30B-0BAB-44B7-B419-55C61FB3D7CB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.6.0 x64:
SET GUID={AF002E58-F25F-4AC2-A360-651F10858F45}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.6.0 Localization package:  
SET GUID={D059EF14-9107-42F6-8F23-EC012E676DF6}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.6.0 x64 Localization package:
SET GUID={E4D83251-3C10-4AD0-A3EE-A4E0B21F5B2D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.6.0:
SET GUID={2DFAF20D-F28A-488F-9867-68A4364F627B}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.6.0 x64:
SET GUID={980A7042-4B0E-4644-AFDC-9A160378149A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.6.0 Localization package:  
SET GUID={B9EF13E7-F977-4379-8813-74D2B687CDBF}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.6.0 x64 Localization package:
SET GUID={D6501BA6-7DF8-4546-9D1A-6D349D090497}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.6.0: (32-bit)
SET GUID={44ABA429-C1F2-4CC4-BB73-3B63BAD041C3}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.6.0: (64-bit)
SET GUID={44ABA429-C1F2-4CC4-BB73-3B63BAD041C3}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM============================================

REM Uninstall Bluebeam Revu 15.5.0: 
SET GUID={2075D48E-90D7-4A57-9BEF-8201E70A7DE0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.5.0 x64:
SET GUID={12E9006D-9A89-4C32-A83B-440039DEAE0E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.5.0 Localization package:  
SET GUID={56C5C7C8-783A-4AFD-A040-CBD6653260B7}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.5.0 x64 Localization package:
SET GUID={7EE9C3AF-F276-4672-9C0A-CC783A99EF83}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.5.0:
SET GUID={73066BDE-5664-49B6-8731-27EC72752FF4}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.5.0 x64:
SET GUID={2915835C-B2F3-424E-9EE3-F3C95FF905E9}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.5.0 Localization package:  
SET GUID={58AE4166-61B7-4805-9AF8-6AA57E14D4EB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.5.0 x64 Localization package:
SET GUID={190D66D6-0B90-4191-9CA3-7781B7F1477A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.5.0: (32-bit) 
SET GUID={02F7CCF9-2DF6-47A6-AB92-279CAFFE206F}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.5.0: (64-bit) 
SET GUID={02F7CCF9-2DF6-47A6-AB92-279CAFFE206F}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================

REM Uninstall Bluebeam Revu 15.3.0: 
SET GUID={A42D742A-51B0-49FD-9CEE-DA1B6F32C97E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.3.0 x64:
SET GUID={23443D56-EA0E-440F-8712-D245CF170348}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.3.0 Localization package:  
SET GUID={C35E4677-4E71-45BE-8FB4-E323C0A3D59C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.3.0 x64 Localization package:
SET GUID={E82AF890-8BAD-439A-827A-237676F78119}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.3.0:
SET GUID={93C4C724-35F0-4268-B4D8-4DB9CCAFE35A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.3.0 x64:
SET GUID={DA7E1DC4-A978-4EFD-B8D8-35879CB76061}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.3.0 Localization package:  
SET GUID={3394C726-31B2-4186-8AC6-A48643D909CD}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.3.0 x64 Localization package:
SET GUID={4650DDF5-E126-46B0-826E-C9E23A9C317E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.3.0: (32-bit)
SET GUID={2682ED0B-8AF1-44A4-97F0-820E93F13322}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.3.0: (64-bit)
SET GUID={2682ED0B-8AF1-44A4-97F0-820E93F13322}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM =========================================

REM Uninstall Bluebeam Revu 15.2.1: 
SET GUID={339893AD-15C4-43F6-9778-C79B367A6CC4}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.2.1 x64:
SET GUID={22C76099-1CD3-4022-8E73-073DC8054DCD}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.2.1 Localization package:  
SET GUID={4BBF10E0-855D-419F-950A-3C3C85F3EA8A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.2.1 x64 Localization package:
SET GUID={A9DC57A8-F549-4FFE-9D88-E919756AF96E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.1:
SET GUID={29B49C73-481B-4A57-9D08-AFD3BA6CC232}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 15.2.1 x64:
SET GUID={F1705B6C-E17A-4F38-8BC0-0F08DA78DB90}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.1 Localization package:  
SET GUID={A890C2E9-E979-4CB5-87CC-4F55A46B00F6}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.1 x64 Localization package:
SET GUID={5DD61512-517B-4846-A5A8-28E41CD30475}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.2.1: (32-bit)
SET GUID={EC2D5913-5053-4CE1-96D7-C24ED73036B9}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.2.1: (64-bit)
SET GUID={EC2D5913-5053-4CE1-96D7-C24ED73036B9}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================

REM Uninstall Bluebeam Revu 15.2.0: 
SET GUID={BDCD35D8-9395-4C17-916D-410CE7BEE7C8}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.2.0 x64:
SET GUID={18B2648C-1121-4733-81EC-9D6901F4E747}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn
 
REM Uninstall Bluebeam Revu 15.2.0 Localization package:  
SET GUID={A400D090-B8F2-4540-8919-A0C4556DBC78}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 15.2.0 x64 Localization package:
SET GUID={1F80DFE9-452A-4F32-A839-122919229D4C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.0:
SET GUID={54318DBB-4DDF-4426-876C-3FC94B8D17AA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.0 x64:
SET GUID={6AC1CB57-CF67-496C-B352-CF8F77F3ADA9}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul 
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.0 Localization package:  
SET GUID={7482AB6D-E26C-4DF2-98BD-AF577E71ED85}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.2.0 x64 Localization package:
SET GUID={A3B02CB6-6C17-4614-AB22-5BF23F5669EC}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.2.0: (32-bit)
SET GUID={3578AA77-8EBE-4D46-BB44-21C75C5E47A1}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.2.0: (64-bit)
SET GUID={3578AA77-8EBE-4D46-BB44-21C75C5E47A1}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================

REM Uninstall Bluebeam Revu 15.1.1: 
SET GUID={11206E68-98D8-4D69-8784-52D50C333C37}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.1.1 x64:
SET GUID={4EB84A42-2F0B-4416-922F-BFAB5FC403CA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.1.1 Localization package:  
SET GUID={C47398AF-17B9-4B2E-95A6-F5665B3A65DB}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.1.1 x64 Localization package:
SET GUID={5BAFEAD4-F5AB-4BAC-A4B8-BCD6DAE21BD6}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.1.1:
SET GUID={75A54071-6172-4B63-89C8-84FBCBE8DE04}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.1.1 x64:
SET GUID={50001C38-BE51-4B13-B1A9-3838FB16BB7D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.1.1 Localization package:  
SET GUID={062B355D-CF11-4307-9304-873409CEC460}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 15.1.1 x64 Localization package:
SET GUID={E82A8F66-739A-4F4C-95FF-40F37E2A30F0}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.1.1: (32-bit)
SET GUID={7C57CC59-96F1-4941-9669-A01C1AD7DA13}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.1.1: (64-bit)
SET GUID={7C57CC59-96F1-4941-9669-A01C1AD7DA13}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================

REM Uninstall Bluebeam Revu 15.1.0: 
SET GUID={F5E42D79-3E4F-4E69-92EC-05D628E70110}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.1.0 x64:
SET GUID={38E296C5-9EE0-4173-9725-71FDF28DC14E}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul 
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.1.0 Localization package:  
SET GUID={9476DA3E-E70D-41B5-A4F8-DE9ED2C09717}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Revu 15.1.0 x64 Localization package:
SET GUID={31FC0814-0AA2-4382-8145-BA61B52C7A7C}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.1.0:
SET GUID={51ABD64A-6218-4267-8035-70B495192828}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.1.0 x64:
SET GUID={5026BE3C-2E8F-409E-8465-40A00DE56FEA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.1.0 Localization package:  
SET GUID={34788CB0-CF3B-4EDF-89BE-62B81EE0DCE5}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn 

REM Uninstall Bluebeam Vu 15.1.0 x64 Localization package:
SET GUID={186222C9-6920-4533-8E3C-71D0807B6885}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.1.0: (32-bit)
SET GUID={192B964E-699D-4A64-A6DD-5727EC351BDE}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.1.0: (64-bit)
SET GUID={192B964E-699D-4A64-A6DD-5727EC351BDE}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================

REM Uninstall Bluebeam Revu 15.0.4: 
SET GUID={2EBD1669-9E89-4683-862D-E5A26C7C1087}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.0.4 x64:
SET GUID={2ADB5465-E0DB-4F3A-BFB0-0F1801E77CE7}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.0.4 Localization package:  
SET GUID={F82F51F1-A8AD-4A88-B7D9-E2FE3BB8C55D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Revu 15.0.4 x64 Localization package:
SET GUID={B2B79E84-6E5C-4065-8551-3A2718BB745D}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.0.4:
SET GUID={C2B54276-BE4F-4459-89ED-BF08D3C33640}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.0.4 x64:
SET GUID={23277C3B-C2ED-4343-8D19-A07632B0FF6A}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.0.4 Localization package:  
SET GUID={7AD259FE-2C4E-44FA-9031-DE3C38B402AA}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam Vu 15.0.4 x64 Localization package:
SET GUID={AC7ADFC0-0648-4C28-AE5F-CE1D927B0AA7}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.0.4: (32-bit)
SET GUID={C7C9EA4C-FBD9-4562-8012-C93A2F490292}
REG QUERY %UNINSTALLKEY%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM Uninstall Bluebeam eXtreme 15.0.4: (64-bit)
SET GUID={C7C9EA4C-FBD9-4562-8012-C93A2F490292}
REG QUERY %UNINSTALLKEY64%\%GUID%>nul 2>nul
IF "%ERRORLEVEL%"=="0" msiexec.exe /x %GUID% /qn

REM =========================================
REM Bluebeam v15 RC
REM =========================================

REM Uninstall Bluebeam Revu 15.0.3: 
SET GUID={00A90C00-B051-4EA3-9022-8C0519D893D1}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.3 x64:
SET GUID={C4A423E3-4B4C-48AD-85F8-A2F1C9E883A8}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.3 Localization package:  
SET GUID={BB23D0B8-EF0E-43FA-8636-E36E481382B3}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.3 x64 Localization package:
SET GUID={3273A461-CEF5-4ABA-BE92-3C6D4176DC8B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.3:
SET GUID={0D9249F5-F578-47D3-9410-CCC21F69CE21}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.3 x64:
SET GUID={3273A461-CEF5-4ABA-BE92-3C6D4176DC8B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.3 Localization package:  
SET GUID={97F2BE23-76EB-44B8-BC32-EA55F6039E17}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.3 x64 Localization package:
SET GUID={20DC0289-BA6A-4745-A650-4C3BB5E116F0}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.3 eXtreme Module:  
SET GUID={84479351-714D-4620-9791-2C580F9709B3}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v15 beta3
REM =========================================

REM Uninstall Bluebeam Revu 15.0.2: 
SET GUID={E890A622-A7A7-4C6D-907B-88226A417A74}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.2 x64:
SET GUID={8CD3B3C7-CBCA-43D9-B989-98A442C14BDF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.2 Localization package:  
SET GUID={C6F1220E-AF6F-4956-8C90-EEA9529FEBC5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.2 x64 Localization package:
SET GUID={EDD52A12-75E4-48FA-A062-6D7F7D469137}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.2:
SET GUID={BB786FCC-53AB-4C36-AE90-CA569FA3C6F5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.2 x64:
SET GUID={53FD89F9-CA90-4CB8-A1E4-E923CDDD27E3}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.2 Localization package:  
SET GUID={44DD34B4-751C-43A3-B4C0-7F4952526119}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.2 x64 Localization package:
SET GUID={0CFE1312-65F9-4716-BE3D-0541CD266014}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.2 eXtreme Module:  
SET GUID={F18B8D1D-A1C1-471C-B0B4-BB2F043C7EE2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v15 beta2
REM =========================================

REM Uninstall Bluebeam Revu 15.0.1: 
SET GUID={406DB480-77CE-4394-9F5F-4A210C793C8D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.1 x64:
SET GUID={B919B0C0-52B3-430D-A50B-A2B051956F36}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.1 Localization package:  
SET GUID={0CB0BC76-8653-4D19-8207-92AFD0FFC460}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.1 x64 Localization package:
SET GUID={E7DE8A83-5E03-4005-AC13-9933C72EF40D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.1:
SET GUID={7061AFA3-4C67-4DCD-8D40-86719BCA6EF3}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.1 x64:
SET GUID={017AAA4C-240D-40B7-948E-BE28398AB4C7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.1 Localization package:  
SET GUID={8D6A4FA4-2A35-486A-BBFD-9C7695AC32F4}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.1 x64 Localization package:
SET GUID={7E7D4118-101A-4DF4-8449-AE9D72170EBC}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.1 eXtreme Module:  
SET GUID={AD82E68B-AE70-462F-A377-14B041D5DEA9}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v15 beta1
REM =========================================

REM Uninstall Bluebeam Revu 15.0.0: 
SET GUID={8C1A9E49-E09A-44E1-A4FB-E9514D738CC5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.0 x64:
SET GUID={D6550563-9C21-49E0-8F07-5A7BE63ED269}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 15.0.0 Localization package:  
SET GUID={CB8A952E-5175-48D2-B048-BD739F90F726}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.0 x64 Localization package:
SET GUID={1867D6A1-83B7-4079-BE5D-76BDC0F85917}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.0:
SET GUID={2A85F8C4-2C41-4157-BFED-4B7105A01AF9}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.0 x64:
SET GUID={62CB6759-97CD-4698-B44F-94C96FB25F37}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 15.0.0 Localization package:  
SET GUID={A9DC2595-601E-4270-BCF7-F3C590392482}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 15.0.0 x64 Localization package:
SET GUID={58AA7F04-F119-4B5A-AC94-831E04F60C2A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 15.0.0 eXtreme Module:  
SET GUID={956C6ABC-496F-492D-85DC-0C9673A0FEF2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v12
REM =========================================

REM Uninstall Bluebeam Revu 12.6.0: 
SET GUID={8C284678-3F62-48F1-8B2C-2B102D2D6867}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.6.0 x64:
SET GUID={CAF3E4B8-B35F-4188-BCEC-34CE2D41323C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.6.0: 
SET GUID={09F59A8D-8BE8-436A-BE96-0436075CF265}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.6.0 x64:
SET GUID={3C648013-9567-4290-BE70-55564B0B0003}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.6.0 eXtreme Module:  
SET GUID={F8407584-C221-4359-9C9B-6BBF842E6675}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.6.0 Localization package:  
SET GUID={FAC9853A-E045-499E-A08A-DAFAA698CA3F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.6.0 x64 Localization package:
SET GUID={BEE2E0B1-CA9B-48D6-9E93-DDA2298A7D8A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.6.0 Localization package:  
SET GUID={5033C5D0-F1C2-4E35-B833-EF99C796DAA1}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.6.0 x64 Localization package:
SET GUID={86CDF91D-C1B8-4FC1-A8F3-8D1210172A8C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.5.0: 
SET GUID={180A6F26-E91D-4F6C-8858-0CA1FCF6FABF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.5.0 x64:
SET GUID={8F81B206-1111-4EFA-8431-42BB992C5D76}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.5.0: 
SET GUID={02BDD6A3-248D-415B-A333-6244CD6D337D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.5.0 x64:
SET GUID={6FC79F2F-D92D-46B5-ABF7-9191429C7123}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.5.0 eXtreme Module:  
SET GUID={E15A3E1F-9066-4B1E-B85F-BC89443B2905}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.5.0 Localization package:  
SET GUID={4A75A63E-7042-46B6-872A-D42296A0F13C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.5.0 x64 Localization package:
SET GUID={D1EA5F81-D8E6-405B-A09F-F61539F1C1EE}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.5.0 Localization package:  
SET GUID={6C547056-CB1E-448F-9EAC-D95F7FB7B037}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.5.0 x64 Localization package:
SET GUID={A06A8B4D-C832-4936-BB9B-F1E6CECDBA45}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.1.0: 
SET GUID={67542532-511E-422C-B7F5-BBFD8C4F0B84}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.1.0 x64:
SET GUID={81D4867E-366F-4F34-A1C2-DF819B7BCF00}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.1.0: 
SET GUID={F46B346E-9FE7-4D73-AC63-19113161EC54}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.1.0 x64:
SET GUID={E8E5EDE8-E5E7-4CC8-9B1C-49A6BF479063}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.1.0 eXtreme Module:  
SET GUID={4DD12426-A1D8-4645-A4A3-E8E09496EB0A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.1.0 Localization package:  
SET GUID={6A57C681-C782-41F1-AD3C-9D6D81DBC729}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.1.0 x64 Localization package:
SET GUID={5A0A5175-E9EB-4BD5-A7AE-25939B2E105E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.1.0 Localization package:  
SET GUID={592C0091-9B25-4682-BF89-CD839961210E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.1.0 x64 Localization package:
SET GUID={4D564400-285D-4E20-A5BE-EB84B09DABE6}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet


REM Uninstall Bluebeam Revu 12.0.1: 
SET GUID={C63AF4A0-91F6-4CF1-957D-8FA9F7400B5B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.0.1 x64:
SET GUID={438E2404-76A0-423D-B076-EE3CB629EFAD}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.0.1 Localization package:  
SET GUID={AA4ADDE6-DC8C-42A6-A5D8-058325C401F1}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.0.1 x64 Localization package:
SET GUID={484D2BEA-B5F1-41CD-9BC2-3D9CE511737B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.0.1: 
SET GUID={3B5758A0-7057-4799-9170-5BDE976D7820}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.0.1 x64:
SET GUID={CBBBF4CD-FE1E-4748-AB9C-13D55E13DAA5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.0.1 Localization package:  
SET GUID={AF1D509B-6457-4F5F-8CB8-A4FABED53EDA}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.0.1 x64 Localization package:
SET GUID={D6FC373E-16DB-4E6F-9307-5A025D707461}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.0.1 eXtreme Module:  
SET GUID={A9B8DF84-DA71-4F85-BC57-2432FEB83A69}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.0.0: 
SET GUID={898B0AC6-283D-40AE-B107-C1C0EC9A9D87}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.0.0 x64:
SET GUID={A8E3F673-82B9-4AF0-97C7-4DEDA7042E5E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 12.0.0 Localization package:  
SET GUID={AF579F0A-11D9-47D6-B76E-C8EAF3B1AB4F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.0.0 x64 Localization package:
SET GUID={68EC0CF2-498E-4852-9352-880757108B64}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.0.0: 
SET GUID={BE38EC2E-E94D-4383-8C87-B418EF76D1D5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.0.0 x64:
SET GUID={6BCBAC34-A993-4511-B06D-2EDA026EB90F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 12.0.0 Localization package:  
SET GUID={9A1180D2-7523-44E4-AA07-94EB13357EE5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 12.0.0 x64 Localization package:
SET GUID={5EDAE0AA-E5D5-4B3E-BC52-7870F7DEA539}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 12.0.0 eXtreme Module:  
SET GUID={E15A3E1F-9066-4B1E-B85F-BC89443B2905}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v11
REM =========================================

REM Uninstall Bluebeam Revu 11.8.0: 
SET GUID={3B203366-253B-4583-BFA8-4E564E14289E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.8.0 x64:
SET GUID={29E6EB6B-B6D5-4B14-87C8-EC359E64108B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.8.0 Localization package:  
SET GUID={C8BA506E-6150-4536-B31A-C2B8CE2149E8}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.8.0 x64 Localization package:
SET GUID={437D90B0-F3DC-44F5-A815-B3A40420E315}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.8.0:
SET GUID={F4F425F2-D15A-427F-A2F8-0EAEE9188207}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.8.0 x64:
SET GUID={D84B2D36-D61B-4BB5-9312-4FD02CCDF2B7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.8.0 Localization package:  
SET GUID={6A3D856C-459F-430D-B057-4BB7E5022125}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.8.0 x64 Localization package:
SET GUID={67BA9681-159B-4D59-94E5-12E5DB17193F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.7.0: 
SET GUID={34F24719-80E5-418B-9375-22A259ABA291}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.7.0 x64:
SET GUID={A62360EF-9FE0-472D-B976-8F3ED8922380}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.7.0 Localization package:  
SET GUID={4936750B-F1D4-4F81-8904-36927C6C20D9}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.7.0 x64 Localization package:
SET GUID={C84181DC-1FAD-40A4-8494-E87E948349B6}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.7.0:
SET GUID={73C8EA0C-F0C2-42B6-AAD6-61C7F2F51F03}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.7.0 x64:
SET GUID={8125EE5C-BA0D-41D3-B69A-83AA84B9CEC5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.7.0 Localization package:  
SET GUID={899718DF-3E32-419D-81CB-F0FC240BE6B9}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.7.0 x64 Localization package:
SET GUID={C3D77A2E-81A1-4A81-B39F-75D2AD7024BD}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.6.0: 
SET GUID={A610E2C5-F820-45EF-B555-A6A2D411D527}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.6.0 x64:
SET GUID={FAC5F00B-0E05-4EA9-A48D-E496296AF75B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.6.0 Localization package:  
SET GUID={7FEF3A65-16FD-4C8C-8F27-535BB1AEFF67}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.6.0 x64 Localization package:
SET GUID={97BCE2BA-3E28-4AC0-A807-3E81B691E57F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.6.0:
SET GUID={EF615371-9A12-4370-86AA-939432D9F141}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.6.0 x64:
SET GUID={DA0383E1-E712-45B5-A7CF-DF00DF023DE7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.6.0 Localization package:  
SET GUID={A001B8BE-309B-4E9F-B1DE-8669C026AA62}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.6.0 x64 Localization package:
SET GUID={943A6126-3067-42A9-9773-63B492ED8A2A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.5.0: 
SET GUID={ABB1FFA1-6EAE-4C6C-B38C-35DB35337B8D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.5.0 x64:
SET GUID={B0E161F2-FA92-42F7-AB71-539B5D3093B2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.5.0 Localization package:  
SET GUID={E3369109-22CA-4766-B599-5DD62305F43F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.5.0 x64 Localization package:
SET GUID={3875DF0D-60EE-4BDC-B27D-967668483509}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.5.0: 
SET GUID={FB4EE68E-B8E1-4C30-91E7-E67642502A3E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.5.0 x64:
SET GUID={66922193-2184-449E-8B17-7058325BD1AD}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.5.0 Localization package:  
SET GUID={6069C0FA-9AF9-4189-9A90-7C3FF7794100}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.5.0 x64 Localization package:
SET GUID={66F07802-C213-47D4-88B6-69FC73C008B4}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.5.0 eXtreme Module:  
SET GUID={99C720D9-B033-44D8-8295-9A71D86A755B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.1.0: 
SET GUID={5D135193-9690-437D-81BB-7D4F067339B7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.1.0 x64:
SET GUID={B8E32CB2-EFFC-46D0-9D20-6B2408863D1F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.1.0 Localization package:  
SET GUID={4F098ACB-301C-402C-ABD4-44C3FB3EBFBE}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.1.0 x64 Localization package:
SET GUID={91BE0BEA-50B9-4191-8A88-7CE37F2A8749}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.1.0: 
SET GUID={DE0B0112-C23D-45D6-898B-81282EE29CB2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.1.0 x64:
SET GUID={263E806C-FD91-44BF-A64C-8396FC8EEDA0}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.1.0 Localization package:  
SET GUID={391AD403-3AC1-4199-9EA3-7F4C8CCD5845}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 11.1.0 x64 Localization package:
SET GUID={E56DCBCC-C9FD-4F42-9B38-8698060DFC4C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 11.0.0: 
SET GUID={2725054A-6EA0-4F8D-9C66-3AF9F81493EF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.0.0 x64:
SET GUID={ACCED714-B4D6-4129-8295-912E962F9B50}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.0.0: 
SET GUID={FECE3B27-ED97-44B1-96C6-493086EA35E5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 11.0.0 x64:
SET GUID={1F8864CB-811A-49D3-AD46-46B7C3B2E3E2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 11.0.0 eXtreme Module:  
SET GUID={1D1DDA26-A438-4169-9FF0-A6AF68A63CF7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v10
REM =========================================

REM Uninstall BluebeamRevu 10.2.3: 
SET GUID={6CCDCA56-9705-42EB-A356-B16B11B76726}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.3 x64:
SET GUID={2BF0E23C-3541-481C-974A-E7F380C057EC}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.3 Localization package:  
SET GUID={FE7A0391-0FEB-48C8-8A28-AFD7787F0827}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.3 x64 Localization package:
SET GUID={602ED4BF-4848-4CA9-896D-A63814683749}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.3: 
SET GUID={9FA7CF50-D3C2-4173-8868-35B04A54BB0F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.3 x64:
SET GUID={C480A4C5-380B-4742-AD09-8892CA2ED294}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.3 Localization package:  
SET GUID={2DDE60A5-2CFB-4782-B904-1B16B5CFD35B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.3 x64 Localization package:
SET GUID={0B82A444-58DA-420D-8BC4-9FB141CD77CA}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.2: 
SET GUID={553C2294-40AF-4AA4-8D36-226B7A35E28D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.2 x64:
SET GUID={3FFE72B9-300E-42C9-9EE1-973A2254828F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.2 Localization package:  
SET GUID={1547C7FD-FAC4-4267-8C1E-E12E98868F32}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.2 x64 Localization package:
SET GUID={CC034FEF-E6BF-4987-B17C-08C6F1F48E39}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.2: 
SET GUID={846C44BC-40C2-4239-9860-382577E71CAB}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.2 x64:
SET GUID={0311424C-3B08-4A39-9AA0-C06242979FCA}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.2 Localization package:  
SET GUID={AE1F321E-FCDA-48C3-8D8F-1FF175AB644C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.2 x64 Localization package:
SET GUID={EB583D3A-3B3B-4A42-B747-A90604E4DB36}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.1: 
SET GUID={1F97432A-7823-4367-90C2-B586246D7BC0}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.1 x64:
SET GUID={08A3001B-2E6D-48D1-AC5C-3E27A7127244}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.1 Localization package:  
SET GUID={84130C37-1070-4085-855D-4A68D5034272}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.1 x64 Localization package:
SET GUID={99C8E2E1-EC8C-478C-98E2-2B8F08A5CE96}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.1: 
SET GUID={1207819E-4BBB-4D7E-9D56-7FE513C24B8B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.1 x64:
SET GUID={3ABC939F-9579-4A71-B161-A4264A2592B8}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.1 Localization package:  
SET GUID={073E4C4A-8F0D-423F-9958-382BB07B93F4}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.1 x64 Localization package:
SET GUID={F5DACAE1-4F0F-4ACE-8AEF-899C5BB7010A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.0: 
SET GUID={A34C4122-C1C1-4211-9366-B141E1813D5C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.0 x64:
SET GUID={1752F2AE-CC29-4F96-9E76-9B49498B1D9A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.2.0 Localization package:  
SET GUID={4C78817C-741F-4EB3-BD5F-C74D4A0D84FF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.2.0 x64 Localization package:  
SET GUID={DF7766DB-6C0C-4B02-A07A-98AB7374E7CF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.0: 
SET GUID={FFEA373B-AD22-4AFD-A3E0-41F93F00B5F0}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.0 x64:
SET GUID={4C78817C-741F-4EB3-BD5F-C74D4A0D84FF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.2.0 Localization package:  
SET GUID={885E6D87-6D4F-49A4-978F-0DAF2CAC6A7D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.2.0 x64 Localization package:
SET GUID={CCE78225-E4A0-4ED9-9A7B-D03F7A9EEA8C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.1.0:
SET GUID={836F3A0A-5086-4C93-AD56-BDC21661EDE8}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.1.0 x64:
SET GUID={8C099C5C-D91D-4BC7-93A5-E3DA0F8EB555}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.1.0 Localization package:  
SET GUID={2F6EE1D4-1539-4B9A-84AC-193F4F1D0AF4}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.1.0 x64 Localization package:  
SET GUID={B5325AC8-7EE9-4869-977D-45F0A172D939}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.1.0:  
SET GUID={4918677D-7297-4E41-B528-FACCC463E732}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.1.0 x64:
SET GUID={11C03820-3116-4EA8-AEEC-5F1748B9F9E4}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.1.0 Localization package:  
SET GUID={BB26772C-513C-4CF1-9ED2-11FA23D31FDD}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.1.0 x64 Localization package:  
SET GUID={0CB3EE2C-D657-4C9B-AA16-1FBFB0092C13}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Revu 10.0.0:  
SET GUID={4D5001E1-DC9F-4CCE-BE21-FB94C5107208}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.0.0 x64:
SET GUID={07267DD8-F5BA-4D95-970C-314133E6D3EC}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Revu 10.0.0 eXtreme Module:  
SET GUID={1F7723E9-7DC6-433F-8EAC-70944AD56A39}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam Vu 10.0.0: 
SET GUID={16C0297A-CDB8-4210-84DB-792911C51766}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam Vu 10.0.0 x64:
SET GUID={E01908C8-3E16-42CD-B2B5-728C9098F5E8}
if exist %WINDIR%\installer\%GUID%  "%INSTALLDIR%\%ADMINPATH%" /uninstall & msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v9
REM =========================================

REM Uninstall Bluebeam 9.5.1:  
SET GUID={3B6D1AC7-3B21-4444-AD47-EB0F4D93578A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.5.1 x64:
SET GUID={AD9D0D19-76BD-4F1F-BC89-B446A1511602}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.5.1 Localization package:  
SET GUID={B966488C-7AE4-46BD-84E6-23241D87C3B0}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.5.1 x64 Localization package:  
SET GUID={54996743-06D9-43F0-ADBE-823415126611}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.5.0:  
SET GUID={5F841DB5-4A6C-4ED3-8DCC-524316D8CF4C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.5.0 x64:
SET GUID={7AF56904-5FDC-4D67-87FE-C21E6659668D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.5.0 Localization package: 
SET GUID={F0D51021-5656-4D52-960A-3BDE58C02B49}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.5.0 x64 Localization package:
SET GUID={09D55BAC-F15B-4552-944C-B47E86127CCC}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.2.1:  
SET GUID={1D4FD859-55F0-4769-B1D7-93C72C1F1910}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.2.1 x64:
SET GUID={72E91880-0A41-462D-83A5-ECD5836CF21B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.2.1 Localization package: 
SET GUID={C223F1CB-521D-4982-8752-74ACA4BAFE48}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.2.1 x64 Localization package: 
SET GUID={C91FE19B-EB16-4055-B7E3-2E71DDB47C2A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.2.1 eXtreme Module: 
SET GUID={907C3238-AEB5-46ED-83B1-319476928729}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.2.0:
SET GUID={FDC83AF0-CE69-41D8-8E07-3722618806DF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.2.0 x64:
SET GUID={5B99426C-5945-40AC-BEAE-94E20062B468}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.2.0 Localization package:  
SET GUID={878DEACC-2641-4C30-8D44-A51E14323875}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.2.0 x64 Localization package:  
SET GUID={4CC77D49-954B-4A74-8DC0-610DF140C93C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.2.0 Bb_SlimDX:
SET GUID={4013548A-F0F1-47C9-9233-C16772EEE03A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.0.0:  
SET GUID={B73668E7-67D8-495D-9A41-3090DCB1848B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.0.0 x64:
SET GUID={D40A3A4B-2982-41A3-9180-28F5DA13F59F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 9.0.0 Bb_SlimDX: 
SET GUID={8D684C47-44A0-43F7-8275-887E6F3C0A21}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM Uninstall Bluebeam 9.0.x eXtreme Module:  
SET GUID={18E367C4-976A-4795-A6CE-231D506687C1}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet

REM =========================================
REM Bluebeam v8 International
REM =========================================

REM Uninstall Bluebeam 8.5.1 Intl: 
SET GUID={F31085F6-BBA9-4E4B-BF9D-8FF51B91EB1C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.5.1 x64 Intl:
SET GUID={B46BB820-2349-4351-BB68-9FF3A372D6FF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM (No 8.5.0 Intl)

REM Uninstall Bluebeam 8.0.1 Intl:
SET GUID={678B1EAB-1BF9-4501-9783-5B210CD5F7E8}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.1 x64 Intl:
SET GUID={327259AE-5641-4B43-802B-30CEA8C3CDC5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.0 Intl:
SET GUID={94F20227-ACA6-4DA0-9752-672EFBE97336}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.0 x64 Intl:
SET GUID={F6A08E38-AE01-40CB-BFDC-FCB352E97599}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v8
REM =========================================

REM Uninstall Bluebeam 8.5.1:
SET GUID={664C2418-8D6D-4005-91F0-622BB87E577A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.5.1 x64:
SET GUID={EC3F9C73-BA30-410E-B3F7-2296A6085B2A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.5.0:
SET GUID={134ACDDB-0779-4621-AC8C-9F04B5405773}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.5.0 x64:
SET GUID={7689BC86-F22D-467D-BED9-EC1626A39D76}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.1:
SET GUID={E90F57DF-C1FE-40D1-B9DB-5D30BB785010}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.1 x64:
SET GUID={1D64B1EB-D66A-4AD9-9D9A-A2A4420D8661}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.0:
SET GUID={F2DC73D5-BD4D-4EC5-BF5F-6E86668BF1B4}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 8.0.0 x64:
SET GUID={6A20A45A-254C-45ED-9586-A6625D7FDE29}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v7 International
REM =========================================

REM Uninstall Bluebeam 7.2.1 Intl: 
SET GUID={A18BF6BA-F2F1-4C94-96E3-4A3561E209FC}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.2.1 x64 Intl:
SET GUID={F18CC2F0-4554-4756-B16D-8766DA34E459}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.2.0 Intl:
SET GUID={236DF8FF-91AF-441C-A2C9-E461A982F825}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.2.0 x64 Intl:
SET GUID={089BDA77-197E-48DD-9152-A543DFC87B3F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v7
REM =========================================

REM Uninstall Bluebeam 7.2.1: 
SET GUID={1E7E6D0E-4A08-4D07-A502-15B6C5D34DEB}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.2.1 x64:
SET GUID={AE9E5C3D-AED5-4EE4-A124-FF8DC71960A7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.2.0:
SET GUID={B93EB699-1835-4A90-9D72-A0323D93068B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.2.0 x64:
SET GUID={CB2654A7-E79E-4818-A192-2C224FA6E198}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.0.1:
SET GUID={B2F46890-8A66-4D3E-940E-5BDC9720A3A2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.0.1 x64:
SET GUID={0633A7C1-4886-458E-8773-B85E46297B33}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.0.0:
SET GUID={8ED151FD-8608-46E2-8F78-7BE9F01C266D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 7.0.0 x64:
SET GUID={1D50EB5D-E5F5-41EF-B43E-8563AAF2E40A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v6
REM =========================================

REM Uninstall Bluebeam 6.5.4:
SET GUID={3E8A0E39-C450-413E-9DD3-4AC2E67F2FA1}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.4 x64:
SET GUID={B2F36211-EA94-4BC4-819C-25913C800E8A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.3:
SET GUID={7E8C7630-9DDD-4315-8FB0-F520D19FCB54}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.3 x64:
SET GUID={0CDF4F31-3A3B-4D09-B9E0-3C978E195DAB}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.2:
SET GUID={7187A868-F72D-4EC8-B6D9-34576C1466D2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.2 x64:
SET GUID={984886C8-4FE0-4AE7-9E4E-CB0668085415}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.1:
SET GUID={E84BBD49-8D38-459B-96E3-D88A7291BC37}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.1 x64:
SET GUID={E600FDC2-BD6B-46BE-A7A0-86879544715E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.0:
SET GUID={B673491D-2909-491D-BB0B-CACD6D79532B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.5.0 x64:
SET GUID={F58E8E3A-F45E-4933-A1AA-DF2DA1B0C801}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.2.0:
SET GUID={4C8F6A88-3C1C-4568-82CA-10E6D3C9C126}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.2.0 x64:
SET GUID={16F3C078-158D-4504-8155-C2EA2D027BE8}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.0.1:
SET GUID={72BEC86B-43D6-4CAC-85D1-1D52686A4DDC}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.0.1 x64:
SET GUID={99AA06E5-8CEB-416C-B777-DBC65B4933FE}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.0.0:
SET GUID={C09431EB-CAF1-4CAD-87F8-A84564C4B17A}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 6.0.0 x64:
SET GUID={7E87BA4B-66E4-4D70-B809-FC7B27CFDC33}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v5
REM =========================================

REM Uninstall Bluebeam 5.5.4:
SET GUID={7C26BDD4-DD8D-41B3-978B-78244D17096D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.3:
SET GUID={FB83E662-9360-4278-B97C-6DE0634454EF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.3 x64:                                           
SET GUID={48C75F13-E10A-40C8-BD0A-451BC2AAD699}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.2:                                              
SET GUID={A4D03111-65C4-4094-9682-E33BD3D89B8C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.2 x64:                                           
SET GUID={FCBDF94A-44F2-4C71-B305-6FE1DE4A2CF2}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.1:                                               
SET GUID={BE578BEC-11C3-4FF9-BCA4-A54E86002C41}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.1 x64:                                           
SET GUID={BDC9B5F7-7681-4185-899E-8AA03981CB21}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.0:
SET GUID={9C0D9089-F069-4954-B4BE-C3C43292D723}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.5.0 x64:                                           
SET GUID={7C640B9F-12B1-4F86-B4FB-99E340F30598}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV64%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII64%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.0.1:
SET GUID={16B64140-0BE7-4D00-A24C-0DA6F3A91982}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 5.0.0:
SET GUID={F9DC9E39-AD80-41A3-A501-EA89A973DEBF}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam v4
REM =========================================

REM Uninstall Bluebeam 4.7.1:
SET GUID={81D049A4-E8C6-49FC-995D-C25181C15C14}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 4.7.0:
SET GUID={EEEAC41D-B3EE-4665-B4A4-174BACB50978}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 4.5.0:
SET GUID={CC7F7247-E35A-4D2D-AF29-E49CFC2ABCE4}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 4.2.0:
SET GUID={79FD5A3D-43CC-49B0-A3D9-5C06FA5D2A16}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 4.0.1:
SET GUID={64FDE6B0-E86A-4A57-A158-6F7A1A343D98}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall Bluebeam 4.0.0:
SET GUID={73AD3CCE-B50A-4967-B3F6-DDB3D3BAB15C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM PbPlus AutoCAD v3
REM =========================================

REM Uninstall PbPlus ACAD 3.5.2:
SET GUID={B2E60F76-D192-421A-AE79-4882859A93CD}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus ACAD 3.5.1:
SET GUID={76A98A8E-D363-4A6F-BBA5-11F38E106A8D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus ACAD 3.5.0:
SET GUID={093786CF-0238-46BB-A348-A59F783523A7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus ACAD 3.1.0:
SET GUID={52D71FE1-7431-4702-BEC9-23D23DAD958D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus ACAD 3.0.3:
SET GUID={0C242D49-183F-44D3-8E98-6B720AE31AD8}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus ACAD 3.0.2:
SET GUID={EE5D0844-01E2-47CF-B673-8935BEE4F563}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM PbPlus Solidworks v3
REM =========================================

REM Uninstall PbPlus SolidWorks 3.5.2:
SET GUID={70C16903-1D24-4070-B939-2CB3249E4002}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus SolidWorks 3.5.1:
SET GUID={EF688472-A56E-44D1-9DFE-BC8ADC889233}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus SolidWorks 3.5.0:
SET GUID={EC35C752-9EAA-4536-836E-D85767005030}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus SolidWorks 3.1.0:
SET GUID={69548A3A-0378-4A27-8625-5DC72CA768F5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus SolidWorks 3.0.3:
SET GUID={CD73EAB1-C83E-4D28-80CC-D7DDFFC72683}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus SolidWorks 3.0.2:
SET GUID={5A3E616E-9835-4868-9A55-28BDC1C775F0}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM PbPlus v3
REM =========================================

REM Uninstall PbPlus ACAD 3.0.1 / SW 2.4.1:
SET GUID={3D7F622F-F566-480B-89E5-AC7B74255EE3}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPlus ACAD 3.0.0 / SW 2.4.0:
SET GUID={B6ACE26C-ABD5-4062-84FB-4CDD02C8B1D5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM PbPDF v2
REM =========================================

REM Uninstall PbPDF 2.2.8:
SET GUID={4E810E24-BFAB-44C3-992E-8D774FD945BA}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 2.2.7:
SET GUID={38BED9D8-F436-4B6C-A3AF-A252B8A9C88C}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 2.2.6:
SET GUID={7F75B265-19A0-4627-9C95-20D82A1A1C8B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 2.2.5:
SET GUID={B4819741-0E19-4E98-855F-A1714841616F}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 2.2.1:
SET GUID={48A0C629-86C4-49F5-8FAA-C40B5D86BA47}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 2.2.0:
SET GUID={AD6F40D1-932B-4693-ACDA-F1B3C60715E7}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 2.1.0:
SET GUID={2CB79F5D-87FF-4DBC-809A-157C4B4976DB}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM PbPDF v1
REM =========================================

REM Uninstall PbPDF 1.5.7:
SET GUID={285E4794-503C-45DD-862B-980DB42A4FE5}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 1.5.6:
SET GUID={389F9AD6-946F-4AE5-9BD9-0B15281D636E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 1.5.5:
SET GUID={B622D2D2-9419-4CEE-8BD2-41F1094302DB}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 1.5.4:
SET GUID={BAF635E0-F308-4D2E-B069-B12ACFA3A89B}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 1.5.3:
SET GUID={B96D7117-C853-4C50-A17E-00B0FD8875EB}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall PbPDF 1.4.0 - 1.5.2:
SET GUID={F1F06E54-EAF7-400A-A6E9-C8C6E62C8746}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam Lite v2
REM =========================================

REM Uninstall BbLite 2.5.2:
SET GUID={1F4B07C0-586E-4257-A42A-FB6A5711820E}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall BbLite 2.5.0:
SET GUID={BBBB3338-C963-4BD4-8989-876E77A2110D}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM Uninstall BbLite 2.1.0:
SET GUID={5525D7DF-037E-4AC7-9844-8E21FF008258}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =========================================
REM Bluebeam Lite v1
REM =========================================

REM Uninstall BbLite 1.6.1:
SET GUID={7F5A0C89-A0AD-4EA9-A0CD-10B5B13EA7F3}
if exist %WINDIR%\installer\%GUID%  msiexec.exe /x %GUID% /quiet && ( reg delete %WCV%_%GUID% /f & reg delete %WCVF% /v C:\Windows\Installer\%GUID% /f & rd /s /q "%ISII%\%GUID%" & rd /s /q "%WGF%\%GUID%" )

REM =====================================================================
REM End Uninstall
REM =====================================================================
