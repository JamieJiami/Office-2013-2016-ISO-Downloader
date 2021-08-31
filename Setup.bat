@echo off >nul
color 17
:edition
title Office Downloader
echo 1.Office 2013 Home and Student Retail
echo 2.Office 2013 Personal Retail (Japanese Only)
echo 3.Office 2013 Home and Business Retail
echo 4.Office 2013 Professional Retail
echo 5.Office Word 2013 Retail
echo 6.Office Excel 2013 Retail
echo 7.Office PowerPoint 2013 Retail
echo 8.Office OneNote 2013 Retail
echo 9.Office Outlook 2013 Retail
echo 10.Office Publisher 2013 Retail
echo 11.Office Access 2013 Retail
echo 12.Office Project 2013 Standard Retail
echo 13.Office Project 2013 Professional Retail
echo 14.Office Visio 2013 Standard Retail
echo 15.Office Visio 2013 Professional Retail
echo 16.Office 2016 Home and Student Retail
echo 17.Office 2016 Personal Retail (Japanese Only)
echo 18.Office 2016 Home and Business Retail
echo 19.Office 2016 Professional Retail
echo 20.Office 2016 Professional Plus Retail
echo 21.Office Word 2016 Retail
echo 22.Office Excel 2016 Retail
echo 23.Office PowerPoint 2016 Retail
echo 24.Office Outlook 2016 Retail
echo 25.Office Publisher 2016 Retail
echo 26.Office Access 2016 Retail
echo 27.Office Project 2016 Standard Retail
echo 28.Office Project 2016 Professional Retail
echo 29.Ofice Visio 2016 Standard Retail
echo 30.Office Visio 2016 Professional Retail
set /P c=Version Seclect--)

if /I "%c%" EQU "1" goto :2013Stu
if /I "%c%" EQU "2" goto :2013Personal
if /I "%c%" EQU "3" goto :2013HomeBusiness
if /I "%c%" EQU "4" goto :2013Pro
if /I "%c%" EQU "5" goto :2013Word
if /I "%c%" EQU "6" goto :2013Excel
if /I "%c%" EQU "7" goto :2013PPT
if /I "%c%" EQU "8" goto :2013OneNote
if /I "%c%" EQU "9" goto :2013Outlook
if /I "%c%" EQU "10" goto :2013Publisher
if /I "%c%" EQU "11" goto :2013Access
if /I "%c%" EQU "12" goto :Project2013Standard
if /I "%c%" EQU "13" goto :Project2013Pro
if /I "%c%" EQU "14" goto :Visio2013Standard
if /I "%c%" EQU "15" goto :Visio2013Pro
if /I "%c%" EQU "16" goto :2016Stu
if /I "%c%" EQU "17" goto :2016Personal
if /I "%c%" EQU "18" goto :2016HomeBusiness
if /I "%c%" EQU "19" goto :2016Pro
if /I "%c%" EQU "20" goto :2016ProPlus
if /I "%c%" EQU "21" goto :2016Word
if /I "%c%" EQU "22" goto :2016Excel
if /I "%c%" EQU "23" goto :2016PPT
if /I "%c%" EQU "24" goto :2016Outlook
if /I "%c%" EQU "25" goto :2016Publisher
if /I "%c%" EQU "26" goto :2016Access
if /I "%c%" EQU "27" goto :Project2016Standard
if /I "%c%" EQU "28" goto :Project2016Pro
if /I "%c%" EQU "29" goto :Visio2016Standard
if /I "%c%" EQU "30" goto :Visio2016Pro
goto edition

:Visio2016Pro

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/VisioProRetail.img"
aria2c.exe %link%
goto done

:Visio2016Standard

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/VisioStdRetail.img"
aria2c.exe %link%
goto done

:Project2016Pro



:Project2016Standard

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/ProjectStdRetail.img"
aria2c.exe %link%
goto done

:2016Access

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/AccessRetail.img"
aria2c.exe %link%
goto done

:2016Publisher

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/PublisherRetail.img"
aria2c.exe %link%
goto done

:2016Outlook

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/OutlookRetail.img"
aria2c.exe %link%
goto done

:2016PPT

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/PowerPointRetail.img"
aria2c.exe %link%
goto done

:2016Excel

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/ExcelRetail.img"
aria2c.exe %link%
goto done

:2016Word

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/WordRetail.img"
aria2c.exe %link%
goto done
:2016ProPlus

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/ProPlusRetail.img"
aria2c.exe %link%
goto done
:2016Pro

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/ProfessionalRetail.img"
aria2c.exe %link%

goto done
:2016HomeBusiness

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/HomeBusinessRetail.img"
aria2c.exe %link%

goto done
:2016Personal

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
aria2c https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/ja-JP/PersonalRetail.img
goto Done

:2016Stu

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/%lang%/HomeStudentRetail.img"
aria2c.exe %link%

:Visio2013Pro

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=VisioProRetail"
aria2c.exe %link%

goto Done

:Visio2013Standard

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=VisioStdRetail"
aria2c.exe %link%

goto Done

:Project2013Pro

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=ProjectProRetail"
aria2c.exe %link%

goto Done

:Project2013Standard

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=ProjectStdRetail"
aria2c.exe %link%

goto Done

:2013Access

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=AccessRetail"
aria2c.exe %link%

goto Done

:2013Publisher

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%W&p3=PublisherRetail"
aria2c.exe %link%

goto Done

:2013Outlook

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=OutlookRetail"
aria2c.exe %link%

goto Done

:2013OneNote

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=OneNoteRetail"
aria2c.exe %link%

goto Done

:2013PPT

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=PowerPointRetail"
aria2c.exe %link%

goto Done

:2013Excel

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=ExcelRetail"
aria2c.exe %link%

goto Done

:2013Word

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=WordRetail"
aria2c.exe %link%

goto Done

:2013Pro

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=ProfessionalRetail"
aria2c.exe %link%

goto Done

:2013HomeBusiness

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=HomeBusinessRetail"
aria2c.exe %link%
goto done


:2013Personal

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
aria2c https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=ja-JP&p3=PersonalRetail
goto Done


:2013Stu

echo You can see Language Code here--)https://reurl.cc/qgy7yD
set /P lang=Type Language Code--)

echo Downloading...
title Downloading
for /f "tokens=3 delims=:. " %%f in ('bitsadmin.exe /CREATE /DOWNLOAD "Download aria2" ^| findstr "Created job"') do set GUID=%%f
bitsadmin /transfer %GUID% /download /priority foreground https://uup.rg-adguard.net/dl/aria2/aria2c_x86.exe "%cd%\aria2c.exe"
set link="https://officeredir.microsoft.com/r/rlidO15C2RMediaDownload?p1=db&p2=%lang%&p3=HomeStudentRetail"
aria2c.exe %link%

goto :Done

:Done
echo Done!
set /P more=Need Download More?(Y/N)--)
if /I "%more%" EQU "Y" goto :edition
if /I "%more%" EQU "N" goto :exit

goto Done

:exit

pause
exit