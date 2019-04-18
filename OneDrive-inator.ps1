# Autor: Piotr Stêpieñ
# Skrypt zosta³ napisany pod Windows 10. W wypapdku innych wersji Windowsa mo¿e nie dzia³aæ poprawnie dlatego sprawdza wersjê Windowsa i w wypadku wersji innej ni¿ 10 wysy³a informacjê zwrotn¹ do u¿ytkownika i ewentualnie Admina.

##### Zmienne modyfikowalne.

$OD_name = "OneDrive - Company_Name" # Nazwa folderu OneDrive. Dla kont s³u¿bowych bêdzie to domyœlnie "OneDrive - nazwa_firmy", dla kont prywatnych "OneDrive". POLE WYMAGANE!
$OD_custom = "False" # Je¿eli œcie¿ka do folderu OneDrive bêdzie inna od domyœlnej (%USERPROFILE%), wpisz j¹ tutaj. Zaleca siê zachowanie domyœlnej œcie¿ki i pozostawienie wartoœci "False".
$OD_folder = "OD_Copy" # Nazwa folderu, w którym bêd¹ przechowywane foldery u¿ytkowników. Ma byæ puste ("") je¿eli maj¹ byæ w folderze g³ównym OneDrive.
$prompt_time = 30 #minut # Je¿eli u¿ytkownik nie ma skonfigurowanego programu OneDrive, skrypt poprosi go o zalogowanie. Okreœl co ile minut skrypt ma powtarzaæ proœbê o zalogowanie. Wartoœæ nie mo¿e znajdowaæ siê cytacie ("przyk³adowa_niepoprawna_wartoœæ")!
$OD_limit = 15 #GB # OneDrive posiada limit maksymalnego rozmiaru plików. Limit ten czasami siê zmienia. NAle¿y tutaj wpisaæ aktualny limit w GB.

$move_desktop = "True" # Wybierz czy ma byæ przeniesiony folder Pulpit. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.
$move_documents = "True" # Wybierz czy ma byæ przeniesiony folder Dokumenty. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.
$move_pictures = "True" # Wybierz czy ma byæ przeniesiony folder Obrazy. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.
$move_video = "True" # Wybierz czy ma byæ przeniesiony folder Filmy. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.
$move_music = "True" # Wybierz czy ma byæ przeniesiony folder Muzyka. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.
$move_favorites = "True" # Wybierz czy ma byæ przeniesiony folder Ulubione. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.
$move_downloads = "False" # Wybierz czy ma byæ przeniesiony folder Pobrane. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ przenoszenie.

$nametest = "True" # Sprawdzanie poprawnoœci nazw (usuwa niedozwolone # oraz % z nazw plików i zastêpuje je s³owami "hash" oraz "prcnt". Proces zajmuje od kilku minut do godziny. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ. 
$sizetest = "True" # Sprawdza czy s¹ pliki wiêksze ni¿ $OD_limit i wypisuje je w pliku "$OD_folderk\brak_synchro.txt" oraz wysy³a to samo info na adres mailowy Admina oraz na do œcie¿ki sieciowej je¿eli te opcje zostan¹ ustawione na "True". Pliki wiêksze ni¿ $OD_limit nie bêd¹ synchronizowane. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ.
$PST_Copy = "True" # Kopiuje pliki PST do oryginalnej lokacji ($Env:userprofile\Documents\Pliki Programu Outlook) aby nie ko¿ystaæ z plików PST umieszczonych na OneDrivie. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ. UWAGA, opcja jest aktywna tylko gdy $move_documents = "True".
$WinCheck = "True" # Program informuje u¿ytkownika i opcjonalnie admina o z³ej wersji Windowsa. Zaleca siê zachowanie wartoœci domyœlnej "True". Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ.

$admin_error_win = "True" # Je¿eli "True" Admin bêdzie dostawaæ maile od u¿ytkowników, którzy maj¹ z³¹ wersj¹ Windowsa. Maile mog¹ byæ blokowane przez antyspam dlatego te informacje bêd¹ równie¿ zapisane na wybranej œcie¿ce sieciowej. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ. U¿ytkownik  dostanie informacje o niezale¿nie od tej opcji. Dzia³a tylko gdy $WinCheck ma wartoœæ "True".
$admin_error_size = "True" # Je¿eli "True" Admin bêdzie dostawaæ maile od u¿ytkowników, którzy maj¹ za du¿e pliki. Maile mog¹ byæ blokowane przez antyspam dlatego te informacje bêd¹ równie¿ zapisane na wybranej œcie¿ce sieciowej. Wpisaæ wartoœæ inn¹ ni¿ True aby pomin¹æ. U¿ytkownik zawsze dostanie informacje o b³êdach.
$admin_success = "True" # Je¿eli "True" program bêdzie generowaæ listê u¿ytkowników, którzy przeœli na OneDrive. Plik nosi nazwê "OneDrive_Migration.txt", znajduje siê w $net_error_lok i zawiera nazwy u¿ytkowników, nazwy komputerów, oraz datê przejœcia.

# Poni¿sze wype³niæ tylko je¿eli $admin_error /$admin_success ma wartoœæ "True".
$net_error_lok = "\\server_address\one_drive\" # Œcie¿ka sieciowa, w kótrej bêd¹ trzymane informacje o b³êdach od u¿ytkowników. W przypadku za du¿ych plików pojawi siê plik tekstowy w folderze "BigFiles" z nazw¹ u¿ytkownia oraz dat¹ zawieraj¹cy listê zbyt du¿ych plików. W wypadku z³ej wersji Windows, u¿ytkownicy wraz z ich wersj¹ Windowsa zostan¹ wypiani w pliku "OneDrive_old_Windows.txt". Pozostawiæ "False" aby pomin¹æ.
$admin_mail = "admin@email.address" # Adres e-mail administratora, który ma dostawaæ ewentualne powiadomienia. Ma byæ puste ("") b¹dŸ fa³szywe aby pomin¹æ.
$company_smtp = "serwer_name.mail.protection.outlook.com" # serwer smtp, przez który nast¹pi wysy³anie. Dzia³aj¹ tylko serwery Direct Send.
$smtp_port = "25" # Port smtp serwera.

##### Zmienna okreœlaj¹ca now¹ lokalizacjê folderów profilowych. NIE ZMIENIAÆ.
if ($OD_custom -ne "False")
{
$lok = "$OD_custom\$OD_name\$OD_folder"
$lokreg = "$OD_custom\$OD_name\$OD_folder"
}
else
{
$lok = "$Env:userprofile\$OD_name\$OD_folder"
$lokreg = "%USERPROFILE%\$OD_name\$OD_folder"
}

##### Zmienna okreœlaj¹ca lokalizacjê w rejestrze folderów u¿ytkowników. Wartoœæ domyœlna jest dla Windowsa 10. NIE ZMIENIAÆ.
$Regkey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

##### Poni¿sze komunikaty mo¿na opcjonalnie modyfikowaæ. Wartoœci domyœlne powinny byæ optymalne.

# Tekst komunikatu logowania siê do programu OneDrive.
$dial1_text = "Proszê zalogowaæ siê do programu OneDrive."
# Tekst komunikatu rozpoczêcia konfiguracji programu OneDrive.
$dial2_text = "Nast¹pi konfiguracja programu OneDrive.`nMo¿e po potrwaæ do kilkunatu minut.`nProszê zapisaæ i zamkn¹æ wszystkie dokumenty a nastêpnie wcisn¹æ przycisk 'OK'."
# Tekst komunikatu zakoñczenia konfiguracji programu OneDrive.
$dial3_text = "Program OneDrive zosta³ skonfigurowany.`nPliki znajduj¹ siê w lokalizacji:`n'$lok'."
# Tekst komunikatu o plikach wiêkszych ni¿ $OD_limit.
$TooBig_text = "Na twoim komputerze znajduj¹ siê pliki zbyt du¿e do synchronizacji.`nIch lista znajduje siê w pliku:`n'$lok\brak_synchro.txt'"
# Tekst komunikatu z³ej wersji Windowsa.
$WinError_text = "Aby skonfigurowaæ program OneDrive niezbêdne jest zainstalowanie Windows 10.`nProszê skontaktowaæ siê z dzia³em IT."

##### Koniec zmiennych modyfikowalnych

##### Program sprawdza wersjê Windowsa

#$WinVer = gwmi win32_operatingsystem | % caption
if ((Get-WmiObject Win32_OperatingSystem).Caption -eq "Microsoft Windows 10 Pro")
{
##### START Konfiguracja

##### Sprawdzenie czy pliki ju¿ s¹ na OneDrive'ie
if ((Get-ItemProperty -Path $Regkey -Name Desktop).Desktop -like "*$lok*")
{
& "$Env:userprofile\AppData\Local\Microsoft\OneDrive\OneDrive.exe" /background
#$dial0 = New-Object -ComObject Wscript.Shell
#$dial0.Popup("Program OneDrive jest ju¿ skonfigurowany.`nPliki znajduj¹ siê w '$lok'",0,"$OD_name")
exit
}   
else
{

  ##### Sprawdzenie czy u¿ytkownik jest zalogowany do OneDrive'a kontem s³u¿bowym
  $path = Test-Path "$Env:userprofile\$OD_name\"
  if ($path -ne "True") 
  {
  ##### Wy³¹czenie programu OneDrive
  #Start-Sleep -s 30
  Stop-Process -processname OneDrive -ErrorAction SilentlyContinue

    ##### Pêtla wywo³uj¹ca logowanie do programu OneDrive
    $prompt_time = $prompt_time*60
    while ($path -ne "True")
    {   
    $dial1 = New-Object -ComObject Wscript.Shell
    $dial1.Popup("$dial1_text",0,"$OD_name")
    & "$Env:userprofile\AppData\Local\Microsoft\OneDrive\OneDrive.exe"
    #exit #usuñ to
    ##### Pauza aby u¿ytkownik móg³ siê zalogowaæ
    Start-Sleep -s "$prompt_time"
    
    ##### Sprawdzenie czy u¿ytkownik jest zalogowany do OneDrive'a kontem s³u¿bowym    
    $path = Test-Path "$Env:userprofile\$OD_name\"
    }
   }

    ##### START konfiguracja
    #for ($I = 1; $I -le 100; $I++ )
    #{Write-Progress -Activity "Search in Progress" -Status "$I% Complete:" -PercentComplete $I;}     
    
    ##### Informacja dla u¿ytkownika  
    $dial2 = New-Object -ComObject Wscript.Shell
    $dial2.Popup("$dial2_text",0,"$OD_name")
 
    ##### Wy³¹czenie programów
    Stop-Process -processname outlook -ErrorAction SilentlyContinue
    Stop-Process -processname lync -ErrorAction SilentlyContinue
    Stop-Process -processname winword -ErrorAction SilentlyContinue
    Stop-Process -processname excel -ErrorAction SilentlyContinue
    Stop-Process -processname powerpnt -ErrorAction SilentlyContinue
    Stop-Process -processname notepad -ErrorAction SilentlyContinue
    Stop-Process -processname PDFXCview -ErrorAction SilentlyContinue
    Stop-Process -processname FoxitReader -ErrorAction SilentlyContinue
    Stop-Process -processname AcroRd32 -ErrorAction SilentlyContinue
    Stop-Process -processname explorer
    #Start-Sleep -s 5
    #Stop-Process -processname mstsc
    
    ##### Tworzenie folderu BACKUP
    New-Item -ItemType directory -Path $lok -ErrorAction SilentlyContinue

    ##### Zamiana znaków niedozwolonych na OneDrive'ie
   if ($nametest -eq "True")
   {
    Write-Host "Sprawdzanie poprawnoœci nazw plików.`nMo¿e to potrwaæ kilka minut."
    Get-ChildItem "$Env:userprofile\Desktop\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 9%"
    Get-ChildItem "$Env:userprofile\Documents\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 18%"
    Get-ChildItem "$Env:userprofile\Pictures\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 27%"
    Get-ChildItem "$Env:userprofile\Videos\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 36%"
    Get-ChildItem "$Env:userprofile\Music\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 45%"
    Get-ChildItem "$Env:userprofile\Favorites\" -Recurse |Rename-Item -NewName {$_.name -replace '#','hash'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 54%"
    Get-ChildItem "$Env:userprofile\Desktop\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 63%"
    Get-ChildItem "$Env:userprofile\Documents\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 72%"
    Get-ChildItem "$Env:userprofile\Pictures\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Get-ChildItem "$Env:userprofile\Videos\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 81%"
    Get-ChildItem "$Env:userprofile\Music\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 99%"
    Get-ChildItem "$Env:userprofile\Favorites\" -Recurse |Rename-Item -NewName {$_.name -replace '%','prcnt'} -ErrorAction SilentlyContinue
    Write-Host "Ukoñczono 100%"

    ##### Wy³¹czenie programów
    Stop-Process -processname outlook -ErrorAction SilentlyContinue
    Stop-Process -processname lync -ErrorAction SilentlyContinue
    Stop-Process -processname winword -ErrorAction SilentlyContinue
    Stop-Process -processname excel -ErrorAction SilentlyContinue
    Stop-Process -processname powerpnt -ErrorAction SilentlyContinue
    Stop-Process -processname notepad -ErrorAction SilentlyContinue
    Stop-Process -processname PDFXCview -ErrorAction SilentlyContinue
    Stop-Process -processname FoxitReader -ErrorAction SilentlyContinue
    Stop-Process -processname AcroRd32 -ErrorAction SilentlyContinue
    Stop-Process -processname AcroRd32 -ErrorAction SilentlyContinue
    #Stop-Process -processname explorer
    #Start-Sleep -s 5
   }
   
   ##### Przenoszenie Pulpitu
   if ($move_desktop -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Desktop" -ErrorAction SilentlyContinue
    #robocopy "$Env:userprofile\Desktop" "$lok\Desktop" /e /xf *
    
    $newPathDesktop = "$lokreg\Desktop"
    #$keyDesktop2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Desktop $newPathDesktop
    set-ItemProperty -path $Regkey -name "{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}" $newPathDesktop
    #set-ItemProperty -path $keyDesktop2 -name Desktop $newPathDesktop

    #robocopy "$Env:userprofile\Desktop" "$lok\"
    #Move-Item "$Env:userprofile\Desktop\*" "$lok\Desktop\" -force
    #Get-ChildItem "$Env:userprofile\Desktop" -Recurse | Move-Item -Destination "$lok\Desktop"
    Get-ChildItem "$Env:userprofile\Desktop\*" -Force | Move-Item -Destination "$lok\Desktop\"

    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboDesk = Get-ChildItem -Path "$Env:userprofile\Desktop" | Measure-Object
    if ($RoboDesk.count -ne "0")
    {
    robocopy "$Env:userprofile\Desktop" "$lok\Desktop" /move /e
    }
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Desktop" -Recurse -Force
    }
   }

   ##### Przenoszenie Dokumentów
   if ($move_documents -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Documents" -ErrorAction SilentlyContinue
    
    $newPathDocuments = "$lokreg\Documents"
    #$keyDocuments2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Personal $newPathDocuments
    set-ItemProperty -path $Regkey -name "{F42EE2D3-909F-4907-8871-4C22FC0BF756}" $newPathDocuments
    #set-ItemProperty -path $keyDocuments2 -name Personal $newPathDocuments

    #Move-Item "$Env:userprofile\Documents\*" "$lok\Documents\" -force
    Get-ChildItem "$Env:userprofile\Documents\*" -Force | Move-Item -Destination "$lok\Documents\"
    #Get-ChildItem "$Env:userprofile\Documents" -Recurse | Move-Item -Destination "$lok\Documents"

    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboDok = Get-ChildItem -Path "$Env:userprofile\Documents" | Measure-Object
    if ($RoboDok.count -ne "0")
    {
    robocopy "$Env:userprofile\Documents" "$lok\Documents" /move /e
    }
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Documents" -Recurse -Force
    }
    # Skopiowanie plików PST do domyœlnej lokalizacji
    if ($PST_Copy -eq "True")
    {
      if (Test-Path "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook")
      {
      New-Item -ItemType directory -Path "$Env:userprofile\Documents\Pliki programu Outlook" -ErrorAction SilentlyContinue
      $maxsize = $OD_limit*1000*1024*1024
      robocopy "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook" "$Env:userprofile\Documents\Pliki programu Outlook" /e /MAX:$maxsize
      Get-ChildItem "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook\*" -Force | Move-Item -Destination "$Env:userprofile\Documents\Pliki programu Outlook\" -ErrorAction SilentlyContinue
      #robocopy "$Env:userprofile\$OD_name\$OD_folder\Documents\Pliki programu Outlook" "$Env:userprofile\Documents\Pliki programu Outlook" /mov /e /MIN:$maxsize
      }
    }
   }

   ##### Przenoszenie Obrazów
   if ($move_pictures -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Pictures" -ErrorAction SilentlyContinue
    
    $newPathPictures = "$lokreg\Pictures"
    #$keyPictures2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name "My Pictures" $newPathPictures
    set-ItemProperty -path $Regkey -name "{0DDD015D-B06C-45D5-8C4C-F59713854639}" $newPathPictures
    #set-ItemProperty -path $keyPictures1 -name "{33E28130-4E1E-4676-835A-98395C3BC3BB}" $newPathPictures
    #set-ItemProperty -path $keyPictures2 -name "My Pictures" $newPathPictures

    #Move-Item "$Env:userprofile\Pictures\*" "$lok\Pictures\" -force
    Get-ChildItem "$Env:userprofile\Pictures\*" -Force | Move-Item -Destination "$lok\Pictures\"
    #Get-ChildItem "$Env:userprofile\Pictures" -Recurse | Move-Item -Destination "$lok\Pictures"
    
    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboPic = Get-ChildItem -Path "$Env:userprofile\Pictures" | Measure-Object
    if ($RoboPic.count -ne "0")
    {
    robocopy "$Env:userprofile\Pictures" "$lok\Pictures" /move /e
    }
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Pictures" -Recurse -Force
    }
   }
    
   ##### Przenoszenie Video
   if ($move_video -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Videos" -ErrorAction SilentlyContinue
     
    $newPathVideos = "$lokreg\Videos"
    #$keyVideos2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name "My Video" $newPathVideos
    set-ItemProperty -path $Regkey -name "{35286A68-3C57-41A1-BBB1-0EAE73D76C95}" $newPathVideos
    #set-ItemProperty -path $keyVideos2 -name "My Video" $newPathVideos

    #Move-Item "$Env:userprofile\Videos\*" "$lok\Videos\" -force
    Get-ChildItem "$Env:userprofile\Videos\*" -Force | Move-Item -Destination "$lok\Videos\"
    #Get-ChildItem "$Env:userprofile\Videos" -Recurse | Move-Item -Destination "$lok\Videos"

    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboVid = Get-ChildItem -Path "$Env:userprofile\Videos" | Measure-Object
    if ($RoboVid.count -ne "0")
    {
    robocopy "$Env:userprofile\Videos" "$lok\Videos" /move /e
    }    
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Videos" -Recurse -Force
    }
   }
        
   ##### Przenoszenie Muzyki
   if ($move_music -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Music" -ErrorAction SilentlyContinue
    
    $newPathMusic = "$lokreg\Music"
    #$keyMusic2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name "My Music" $newPathMusic
    set-ItemProperty -path $Regkey -name "{A0C69A99-21C8-4671-8703-7934162FCF1D}" $newPathMusic  
    #set-ItemProperty -path $keyMusic2 -name "My Music" $newPathMusic
    
    #Move-Item "$Env:userprofile\Music\*" "$lok\Music\" -force
    Get-ChildItem "$Env:userprofile\Music\*" -Force | Move-Item -Destination "$lok\Music\"
    #Get-ChildItem "$Env:userprofile\Music" -Recurse | Move-Item -Destination "$lok\Music"

    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboMuz = Get-ChildItem -Path "$Env:userprofile\Music" | Measure-Object
    if ($RoboMuz.count -ne "0")
    {
    robocopy "$Env:userprofile\Music" "$lok\Music" /move /e
    }  
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Music" -Recurse -Force
    }
   }
    
   ##### Przenoszenie Ulubionych
   if ($move_favorites -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Favorites" -ErrorAction SilentlyContinue
    
    $newPathFavorites = "$lokreg\Favorites"
    #$keyFavorites2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Favorites $newPathFavorites  
    #set-ItemProperty -path $keyFavorites2 -name Favorites $newPathFavorites

    #Move-Item "$Env:userprofile\Favorites\*" "$lok\Favorites\" -force
    Get-ChildItem "$Env:userprofile\Favorites\*" -Force | Move-Item -Destination "$lok\Favorites\"
    #Get-ChildItem "$Env:userprofile\Favorites" -Recurse | Move-Item -Destination "$lok\Favorites"

    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboFav = Get-ChildItem -Path "$Env:userprofile\Favorites" | Measure-Object
    if ($RoboFav.count -ne "0")
    {
    robocopy "$Env:userprofile\Favorites" "$lok\Favorites" /move /e
    } 
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Favorites" -Recurse -Force
    }
   }

   ##### Przenoszenie Pobranych
   if ($move_downloads -eq "True")
   {
    New-Item -ItemType directory -Path "$lok\Downloads" -ErrorAction SilentlyContinue
    
    $newPathDownloads = "$lokreg\Downloads"
    #$keyDownloads2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"  
    set-ItemProperty -path $Regkey -name Downloads $newPathDownloads
    set-ItemProperty -path $Regkey -name "{374DE290-123F-4565-9164-39C4925E467B}" $newPathDownloads
    #set-ItemProperty -path $keyDownloads2 -name Downloads $newPathDownloads

    Get-ChildItem "$Env:userprofile\Downloads\*" -Force | Move-Item -Destination "$lok\Downloads\"

    ##### Sprawdzenie czy wszystko zosta³o przeniesione
    $RoboDow = Get-ChildItem -Path "$Env:userprofile\Downloads" | Measure-Object
    if ($RoboDow.count -ne "0")
    {
    robocopy "$Env:userprofile\Downloads" "$lok\Downloads" /move /e
    } 
    ##### Usuniêcie pustych folderów
    else
    {
    Remove-Item "$Env:userprofile\Downloads" -Recurse -Force
    }
   }

   ##### Sprawdzenie plików wiêkszych ni¿ $OD_limit
   if ($sizetest -eq "True")
   {
    $OD_limit = "$OD_limit"+"GB"
    $BigFiles = Get-ChildItem "$Env:userprofile\" -Recurse | Where-Object {$_.Length -gt "$OD_limit"}

   if ($BigFiles) 
   {    
    $date = (Get-Date).ToShortDateString()
    $BigFiles >> "$lok\brak_synchro.txt"
    #$BigFiles | Set-Content "$lok\brak_synchro.txt" #same nazwy, bez œcie¿ek
    
    ##### Wys³anie maila z list¹ plików i nazw¹ u¿ytkownika do admina oraz zapisanie tej informacji na œcie¿ce sieciowej.
    if ($admin_error_size -eq "True")
    {
    Send-MailMessage -To "$admin_mail" -From "$admin_mail" -Subject "OneDrive - $env:UserName - za duze pliki" -Body "$env:UserName`n$BigFiles" -SmtpServer "$company_smtp" -Port "$smtp_port" -Attachments "$lok\brak_synchro.txt" -ErrorAction SilentlyContinue
      if (Test-Path "$net_error_lok")
      {
      $date = (Get-Date).ToShortDateString()
      New-Item -ItemType directory -Path "$net_error_lok\BigFiles" -ErrorAction SilentlyContinue
      $BigFiles >> "$net_error_lok\BigFiles\$date;$env:UserName.txt"
      }
    }
    $TooBig = New-Object -ComObject Wscript.Shell
    $TooBig.Popup("$TooBig_text",0,"$OD_name")
   }
   }
   
    ##### Zapisanie w $net_error_lok\OneDrive_Migration.txt daty, nazwy u¿ytkownika oraz komputera, któremu uda³o siê przenieœæ pliki na OneDrive'a.
    if ($admin_success -eq "True")
    {
     $date = (Get-Date).ToShortDateString()
     New-Item -ItemType directory -Path "$net_error_lok" -ErrorAction SilentlyContinue
     "$date;$env:computername;$env:UserName" >> "$net_error_lok\OneDrive_Migration.txt"
    }
    ##### Restart Explorera i komunikat koñca
    Stop-Process -processname explorer
    #Start-Process explorer
    $dial3 = New-Object -ComObject Wscript.Shell
    $dial3.Popup("$dial3_text",0,"$OD_name")
exit
}
##### END konfiguracja 
}

elseif ($WinCheck = "True")
##### Program informuje o z³ej wersji Windowsa
{
 if ($admin_error_win -eq "True")
 {
  $winpath = Test-Path "$Env:userprofile\WinVer.dat"
  if ($winpath -ne "True")
  {
    if (Test-Path "$net_error_lok")
    {
    $date = (Get-Date).ToShortDateString()
    New-Item -ItemType directory -Path "$net_error_lok" -ErrorAction SilentlyContinue
    "$date;$env:computername;$env:UserName;$WinVer" >> "$net_error_lok\OneDrive_old_Windows.txt"
    }
  Send-MailMessage -To "$admin_mail" -From "$admin_mail" -Subject "OneDrive - $env:UserName - brak Windowsa 10" -Body "$env:UserName`n$WinVer" -SmtpServer "$company_smtp" -Port "$smtp_port" -ErrorAction SilentlyContinue
  "" >> "$Env:userprofile\WinVer.dat"
  }
 }
 $WinError = New-Object -ComObject Wscript.Shell
 $WinError.Popup("$WinError_text",0,"$OD_name")
 exit
}

else {exit}
