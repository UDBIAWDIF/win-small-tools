cls
@echo off

echo ϵͳ��������
echo ����IE,����ȵ�
echo Loading...
del /f /s /q %systemdrive%\*.tmp
del /f /s /q %systemdrive%\*._mp
del /f /s /q %systemdrive%\*.log
del /f /s /q %systemdrive%\*.gid
del /f /s /q %systemdrive%\*.chk
del /f /s /q %systemdrive%\*.old
del /f /s /q %systemdrive%\*.syd
del /f /s /q %systemdrive%\recycled\*.*
del /f /s /q %windir%\*.bak
echo ����IE�ļ�,����ȵ�
del /f /s /q %windir%\prefetch\*.*
rd /s /q %windir%\temp & md %windir%\temp   
del /f /s /q %windir%\temp\*.*
del /f /s /q "%appdata%\Microsoft\Windows\cookies\*.*"
del /f /q %userprofile%\recent\*.*
del /f /s /q "%userprofile%\Local Settings\Temporary Internet Files\*.*"
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows\Temporary Internet Files\*.*"
del /f /s /q  "%userprofile%\Local Settings\Temp\*.*"
del /f /s /q "%userprofile%\recent\*.*"
del /f /s /q %windir%\SoftwareDistribution\Download\*.*
del /f /s /q "%appdata%\Microsoft\Windows\Recent\*.*"
rd /s /q "%userprofile%\Local Settings\Temp\" & md "%userprofile%\Local Settings\Temp\"
del /f /s /q "%userprofile%\Application Data\PPStream\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\PPStream\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\TheWorld\data\*.*"
del /f /s /q "%userprofile%\Application Data\PPLive\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\PPLive\*.*"
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows\Explorer\*.*"
del /f /s /q "%windir%\Offline Web Pages\*.*
del /f /s /q "%userprofile%\UpdateData\*.*"
del /f /s /q "%userprofile%\Application Data\QQUpdate\*.*"
del /f /s /q "%ProgFile%\Thunder Network\Thunder\Program\Update\*.*"
rd /s /q "%ProgramFiles%\Tencent\QQGame\Download"
del /a /f /s /q "%ProgramFiles%\Kaspersky Lab\*.tmp"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Recorded TV\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Pictures\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Music\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Videos\*.*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\PPLive\*.*"
rd /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\thunder_vod_cache\" 
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\thunder_dctemp\*.*" 
del /f /s /q "%systemdrive%\Documents and Settings\Application Data\Kaspersky Lab\AVP8\Data\*.*"
del /a /f /s /q "%HOMEDRIVE%\pragramdata\All Users\Application Data\Kaspersky Lab\AVP8\*.*"
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows Photo Gallery\Original Images\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\Kingsoft\KavRep\*.*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Kingsoft\KavRep\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\Kingsoft\office6\backup\*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Search\*.*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Windows Defender\Backup\*.*"
del /f /s /q "%ProgramFiles%\Microsoft Security Essentials\Backup\*.*"
del /a /f /s /q "%userprofile%\Application Data\SogouPY\sgim_ext.bin"
del /f /s /q "%userprofile%\AppData\LocalLow\SogouPY\sgim_ext.bin"
del /a /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Microsoft Antimalware\LocalCopy\*.*"  
del /a /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Windows Defender\LocalCopy\*.*" 
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Thunder Network\Thunder_A30B0AF7-D81B-464e-B4E4-4B6DF996FB46_\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Thunder Network\thunder_DEF2CF18-A4A5-468A-ADE3-B410C49964BC_\Temp\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Thunder Network\Thunder_xmp\*.*"

del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\���������Ӱ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\�������ֺ�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\UUSee �������.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\��鿴2.0.*"

del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\�������վ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\��ɽ����ר��2009.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\��ɽ����.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\Ѹ�׿���-��Ѹ���Ӱ��.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\ǧǧ����.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\������ʾ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\������Ϸ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\�������վ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\�޸�360��ȫ��ʿ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\������Ϸ.*"

del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����ʼ���˵�\����\PPLive\�������Ӱ��չ��.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\PPLive\PPLive ��ҳ.*"
del /a /f /q "%ProgramFiles%\PPLiveVA\PPLiveVA.url"
del /a /f /q "%ProgramFiles%\PPLive\PPLive.url"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\360��ȫ��ʿ\�޸�360��ȫ��ʿ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����ʼ���˵�\����\����Ӱ����\����Ӱ���� ��ҳ.*"

del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\PPS�������\����֪�� ���Դ���.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\PPS�������\����������°�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\�ṷ\����֪�� ���Դ���.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\�ṷ\����������°�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\Microsoft Office\���������߾�Ʒ����վ.*"

del /a /f /q "D:\�ҵ��ĵ�\�ղؼ�\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\QQ��ַ��ȫ.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\����\QQ��ַ��ȫ.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\������Թ�����-Jinhu Computer Studio.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\������ - Search All.*"

del /a /f /q "%USERPROFILE%\�ղؼ�\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "%USERPROFILE%\Favorites\QQ��ַ��ȫ.*"
del /a /f /q "%USERPROFILE%\Favorites\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "%USERPROFILE%\Favorites\����\QQ��ַ��ȫ.*"
del /a /f /q "%USERPROFILE%\Favorites\������Թ�����-Jinhu Computer Studio.*"
del /a /f /q "%USERPROFILE%\Favorites\������ - Search All.*"

del /a /f /q "%ProgramFiles%\FlashGet Network\FlashGet 3\��ɫ���ذ�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\�쳵 (FlashGet)\���� ���ɷ�����ɫ���ذɣ�.*"

del /a /f /q "%ProgramFiles%\StormII\��ɫ���ذ�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\����Ӱ��\���� ���ɷ�����ɫ���ذɣ�.*"

del /a /f /q "%SystemDrive%\WINDOWS\Tasks\SogouImeMgr.job"
exit
������Ĵ��븴��ճ�����½��ı����棬�޸��ı���׺����.txt��Ϊ��.Bat��
cls
@echo off

echo ϵͳ��������
echo ����IE,����ȵ�
echo Loading...
del /f /s /q %systemdrive%\*.tmp
del /f /s /q %systemdrive%\*._mp
del /f /s /q %systemdrive%\*.log
del /f /s /q %systemdrive%\*.gid
del /f /s /q %systemdrive%\*.chk
del /f /s /q %systemdrive%\*.old
del /f /s /q %systemdrive%\*.syd
del /f /s /q %systemdrive%\recycled\*.*
del /f /s /q %windir%\*.bak
echo ����IE�ļ�,����ȵ�
del /f /s /q %windir%\prefetch\*.*
rd /s /q %windir%\temp & md %windir%\temp   
del /f /s /q %windir%\temp\*.*
del /f /s /q "%appdata%\Microsoft\Windows\cookies\*.*"
del /f /q %userprofile%\recent\*.*
del /f /s /q "%userprofile%\Local Settings\Temporary Internet Files\*.*"
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows\Temporary Internet Files\*.*"
del /f /s /q  "%userprofile%\Local Settings\Temp\*.*"
del /f /s /q "%userprofile%\recent\*.*"
del /f /s /q %windir%\SoftwareDistribution\Download\*.*
del /f /s /q "%appdata%\Microsoft\Windows\Recent\*.*"
rd /s /q "%userprofile%\Local Settings\Temp\" & md "%userprofile%\Local Settings\Temp\"
del /f /s /q "%userprofile%\Application Data\PPStream\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\PPStream\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\TheWorld\data\*.*"
del /f /s /q "%userprofile%\Application Data\PPLive\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\PPLive\*.*"
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows\Explorer\*.*"
del /f /s /q "%windir%\Offline Web Pages\*.*
del /f /s /q "%userprofile%\UpdateData\*.*"
del /f /s /q "%userprofile%\Application Data\QQUpdate\*.*"
del /f /s /q "%ProgFile%\Thunder Network\Thunder\Program\Update\*.*"
rd /s /q "%ProgramFiles%\Tencent\QQGame\Download"
del /a /f /s /q "%ProgramFiles%\Kaspersky Lab\*.tmp"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Recorded TV\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Pictures\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Music\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Videos\*.*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\PPLive\*.*"
rd /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\thunder_vod_cache\" 
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\thunder_dctemp\*.*" 
del /f /s /q "%systemdrive%\Documents and Settings\Application Data\Kaspersky Lab\AVP8\Data\*.*"
del /a /f /s /q "%HOMEDRIVE%\pragramdata\All Users\Application Data\Kaspersky Lab\AVP8\*.*"
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows Photo Gallery\Original Images\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\Kingsoft\KavRep\*.*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Kingsoft\KavRep\*.*"
del /f /s /q "%userprofile%\AppData\Roaming\Kingsoft\office6\backup\*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Search\*.*"
del /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Windows Defender\Backup\*.*"
del /f /s /q "%ProgramFiles%\Microsoft Security Essentials\Backup\*.*"
del /a /f /s /q "%userprofile%\Application Data\SogouPY\sgim_ext.bin"
del /f /s /q "%userprofile%\AppData\LocalLow\SogouPY\sgim_ext.bin"
del /a /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Microsoft Antimalware\LocalCopy\*.*"  
del /a /f /s /q "%systemdrive%\Documents and Settings\All Users\Application Data\Microsoft\Windows Defender\LocalCopy\*.*" 
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Thunder Network\Thunder_A30B0AF7-D81B-464e-B4E4-4B6DF996FB46_\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Thunder Network\thunder_DEF2CF18-A4A5-468A-ADE3-B410C49964BC_\Temp\*.*"
del /a /f /s /q "%HOMEDRIVE%\Users\Public\Thunder Network\Thunder_xmp\*.*"

del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\���������Ӱ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\�������ֺ�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\UUSee �������.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\Application Data\Microsoft\Internet Explorer\Quick Launch\��鿴2.0.*"

del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\�������վ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\��ɽ����ר��2009.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\��ɽ����.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\Ѹ�׿���-��Ѹ���Ӱ��.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\ǧǧ����.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\������ʾ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\������Ϸ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\�������վ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����\�޸�360��ȫ��ʿ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����\������Ϸ.*"

del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����ʼ���˵�\����\PPLive\�������Ӱ��չ��.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\PPLive\PPLive ��ҳ.*"
del /a /f /q "%ProgramFiles%\PPLiveVA\PPLiveVA.url"
del /a /f /q "%ProgramFiles%\PPLive\PPLive.url"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\360��ȫ��ʿ\�޸�360��ȫ��ʿ.*"
del /a /f /q "%SystemDrive%\Documents and Settings\All Users\����ʼ���˵�\����\����Ӱ����\����Ӱ���� ��ҳ.*"

del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\PPS�������\����֪�� ���Դ���.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\PPS�������\����������°�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\�ṷ\����֪�� ���Դ���.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\�ṷ\����������°�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\Microsoft Office\���������߾�Ʒ����վ.*"

del /a /f /q "D:\�ҵ��ĵ�\�ղؼ�\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\QQ��ַ��ȫ.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\����\QQ��ַ��ȫ.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\������Թ�����-Jinhu Computer Studio.*"
del /a /f /q "D:\�ҵ��ĵ�\Favorites\������ - Search All.*"

del /a /f /q "%USERPROFILE%\�ղؼ�\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "%USERPROFILE%\Favorites\QQ��ַ��ȫ.*"
del /a /f /q "%USERPROFILE%\Favorites\���ʰ�Ӣ��ѧϰ����.*"
del /a /f /q "%USERPROFILE%\Favorites\����\QQ��ַ��ȫ.*"
del /a /f /q "%USERPROFILE%\Favorites\������Թ�����-Jinhu Computer Studio.*"
del /a /f /q "%USERPROFILE%\Favorites\������ - Search All.*"

del /a /f /q "%ProgramFiles%\FlashGet Network\FlashGet 3\��ɫ���ذ�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\�쳵 (FlashGet)\���� ���ɷ�����ɫ���ذɣ�.*"

del /a /f /q "%ProgramFiles%\StormII\��ɫ���ذ�.*"
del /a /f /q "%SystemDrive%\Documents and Settings\Administrator\����ʼ���˵�\����\����Ӱ��\���� ���ɷ�����ɫ���ذɣ�.*"

del /a /f /q "%SystemDrive%\WINDOWS\Tasks\SogouImeMgr.job"
exit