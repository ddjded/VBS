' Префикс названий групп
prefix_1c = "G-1C"

' Получаем имя пользователя
set info = CreateObject( "ADSystemInfo" )

' Получаем учетную запись
set user = GetObject( "LDAP://" & info.UserName )

' Создаем файловые потоки
set res_83 = CreateObject( "ADODB.Stream" )
res_83.Type = 2
res_83.Charset = "UTF-8"
res_83.Open
res_83.Position = 0

set res_82 = CreateObject( "ADODB.Stream" )
res_82.Type = 2
res_82.Charset = "UTF-8"
res_82.Open
res_82.Position = 0

memberOf = user.memberOf

' Просматривает список групп
If (not (IsEmpty(memberOf)) ) then
  For Each item in user.memberOf
    set group = GetObject( "LDAP://" & item )

    if (InStr( group.CN, prefix_1c ) = 1) then
      if (InStr( group.info, "Version=8.3" ) > 0) then
        res_81.WriteText( "[" & group.Description & "]" & Chr(13) & Chr(10) )
        res_81.WriteText( group.info & Chr(13) & Chr(10) )
      end if

      if (InStr( group.info, "Version=8.2" ) > 0) then
        res_82.WriteText( "[" & group.Description & "]" & Chr(13) & Chr(10) )
        res_82.WriteText( group.info & Chr(13) & Chr(10) )
      end if
    end if
  next
end if

' Ищем путь до файлов
set shell = CreateObject( "WScript.Shell" )
appdata = shell.ExpandEnvironmentStrings( "%APPDATA%" )

' Создать папки
set fso = CreateObject( "Scripting.FileSystemObject" )

if (not fso.FolderExists( appdata + "\1C" )) then
  fso.CreateFolder( appdata + "\1C" )
end if
'if (not fso.FolderExists( appdata + "\1C\1Cv81" )) then
'  fso.CreateFolder( appdata + "\1C\1Cv81" )
'end if
if (not fso.FolderExists( appdata + "\1C\1CEStart" )) then
  fso.CreateFolder( appdata + "\1C\1CEStart" )
end if

' И пишем файлы туда

res_83.SaveToFile appdata & "\1C\1CEStart\ibases.v8i", 2
res_83.Close

res_82.SaveToFile appdata & "\1C\1CEStart\ibases.v8i", 2
res_82.Close
