param(
    [Parameter(Mandatory="true",Position=0)]
    $dbInfo,
    [Parameter(Mandatory="true",Position=1)]
    $creds
)

$BaseStatus = @{
    Waiting = "WaitingForUpdate";
    Updating = "UpdateInProgress";
    Disabled = "UpdateDisabled";
    Done = "Updated"
}

#Путь к модулю для подключения к БД, в которой хранится информация о базах 1С
$dbModulePath = "\\bit-live.ru\Scripts\Modules\Auto.psm1"

#Путь к папке с обновлениями 1с
$updates1CRootPath = "\\fscg01ca\Updates"

#Путь к xml-файлу, содержащему информацию о соответствии реальных названий конфигураций 
# и папок с обновлениями
$urlbasesXmlPath = "\\bit-live.ru\scripts\AutoUpdate\conf\URLBASES.xml"
    
#region Functions

    $logFilePath = "\\bit-live.ru\scripts\AutoUpdate\log\1CUpdateManager_V3.txt"
    $logFileSizeInBytes = "5000000"
    $deleteOldLogsAfterDays = 60

function writeLog($message, [switch]$error) {
    if (Test-Path "$logFilePath") {
        $currentLogFile = Get-Item "$logFilePath"
        if ($currentLogFile.Length -ge $logFileSizeInBytes) {
            gci "$($currentLogFile.PSParentPath)" -File | ? {$_.Name -match "$($currentLogFile.Name)$"} | % {
                $now = Get-Date
                $old = $_.LastWriteTime 
                $daysOld = ($now - $old).TotalDays
                if ($daysOld -gt $deleteOldLogsAfterDays) {
                    Remove-Item $_.FullName | Out-Null
                }
            }

            $timestamp = Get-Date -Format "dd-MM-yyyy__hh-mm"
            $newName = $timestamp + "__" + $($currentLogFile.Name)
            Rename-Item -path "$logFilePath" -NewName $newName
        }
    }

    $message = "$(get-date) --> $($env:COMPUTERNAME) --> $($env:username) --> $($dbInfo.basename) --> " + $message
    if($error) {
        Write-Host "$message" -ForegroundColor Red
    } else {
        Write-Host "$message" -ForegroundColor Green
    }
    Write-Output $message | Out-File -FilePath "$logFilePath" -Encoding utf8 -Append -Force
}

    # Получение информации о файловой базе
    function getFileBaseInformation ($dbInfo, $urlbasesXmlPath) {
        try {
            [xml]$urlbasesXml = Get-Content -Path "$urlbasesXmlPath" -Encoding UTF8

            $baseUser = $dbInfo.Login
            if($baseuser -eq $null){$baseUser = ""}
	        $basePassword = $dbInfo.Pass
            if ($basePassword -eq $null){$basePassword = ""}
	        $pathToBase = $dbInfo.Path

	        $comConnector1cObj = New-Object -COMObject "$($Global:connectorName)"
	        $comConnector1C=$ComConnector1cObj.connect("File=""$pathToBase""; Usr=""$baseUser"";Pwd = ""$basePassword"";")
	        if (!($comConnector1C)){             
                throw "не удалось подключиться с помощью COM-коннектора"
	        }
	
            $metadata=[System.__ComObject].InvokeMember("Метаданные",[System.Reflection.BindingFlags]::GetProperty,$null,$comConnector1C,$null)
	        $baseVersion=[System.__ComObject].InvokeMember("Версия",[System.Reflection.BindingFlags]::GetProperty,$null,$metadata,$null)
            $dbInfo.Version = $baseVersion    
            $realBaseConf=[System.__ComObject].InvokeMember("Синоним",[System.Reflection.BindingFlags]::GetProperty,$null,$metadata,$null)
	        $baseConf = ($urlbasesXml.settings.Config | ? {$_.realname -ieq "$realBaseConf"}).name
            $dbInfo.Conf = $baseConf        
            

            writeLog "Проверка наличия соединений с базой"
            $count = countBaseConnections -comConnector1C $comConnector1C
            if ($count -gt 1) {
                throw "Обнаружено $count подключений к базе. Обновление не будет запущено"
            }
            
            $dbInfo            

        } catch {
            if ($error[0] -match "Неправильное имя или пароль пользователя") {
		        $logUser_Password = "\\bit-live.ru\scripts\AutoUpdate\log\logUser_Password.txt"
		        $DisableName = $dbInfo.basename		        
		        Write-Output "$DisableName не правильные логопасы, снимаем с автообновлений" | Out-File -FilePath $logUser_Password -Encoding utf8 -Append
	            $Query = "UPDATE [AutoUpdate].[dbo].[AutoUp] SET Status = 'UpdateDisabled' WHERE BaseName = '$DisableName'"
                sql_Auto -Query $Query | out-null
                throw "$DisableName --> не правильные логопасы, снимаем с автообновлений"
		    }
            elseif ($error[0] -match "Отсутствует файл базы данных") {
		        $logNotPath = "\\bit-live.ru\scripts\AutoUpdate\log\NotPath.txt"
		        $DisableName = $dbInfo.basename		            
		        Write-Output "$DisableName Отсутствует файл базы данных CD, снимаем с автообновлений" | Out-File -FilePath $logNotPath -Encoding utf8 -Append
	            $Query = "UPDATE [AutoUpdate].[dbo].[AutoUp] SET Status = 'UpdateDisabled' WHERE BaseName = '$DisableName'"
                sql_Auto -Query $Query | out-null
                throw "$DisableName Отсутствует файл базы данных CD, снимаем с автообновлений"
		    }
            throw $error[0]
        }
    } 

    # Проверка необходимости обновления
    function checkIsUpdateRequired($dbInfo) {
        writeLog "Определение версии следующего обновления" 
            $nextUpdateVersion = getTrueUpdate -dbInfo $dbInfo -updates1CRootPath $updates1CRootPath
            if ($nextUpdateVersion -ne $null) {
                writeLog "Версия найдена: $nextUpdateVersion" 
                $dbInfo.Status = $BaseStatus.Waiting
                $dbInfo.NextUpdateVersion = $nextUpdateVersion
                writeInfoToDB -dbInfo $dbInfo
                $dbInfo
            } else {
                $dbInfo.NextUpdateVersion = ""
                $dbInfo.Status = $BaseStatus.Done
                writeInfoToDB -dbInfo $dbInfo
                throw "Обновление не требуется либо не найдена следующая версия"
            }
    }

    # Получение информации о базе SQL
    function getSqlBaseInformation ($dbInfo, $urlbasesXmlPath) {
        try {
            [xml]$urlbasesXml = Get-Content -Path "$urlbasesXmlPath" -Encoding UTF8

            $baseUser = $dbInfo.Login
            if($baseuser -eq $null){$baseUser = ""}
	        $basePassword = $dbInfo.Pass
            if ($basePassword -eq $null){$basePassword = ""}
	        $serverName = $dbInfo.sqlserver
            $baseName = $dbInfo.BaseName

            $rmngr = $serverName +":1541"
            $comConnector1cObj = New-Object -COMObject "$($Global:connectorName)"
            $comConnector1C = $ComConnector1cObj.connect("Srvr=""$rmngr""; Ref=""$baseName""; Usr=""$baseUser"";Pwd =""$basePassword"";")
            if (!($comConnector1C)){             
                throw "не удалось подключиться с помощью COM-коннектора"
	        }
            
            $metadata=[System.__ComObject].InvokeMember("Метаданные",[System.Reflection.BindingFlags]::GetProperty,$null,$comConnector1C,$null)
	        $baseVersion=[System.__ComObject].InvokeMember("Версия",[System.Reflection.BindingFlags]::GetProperty,$null,$metadata,$null)
            $dbInfo.Version = $baseVersion    
            $realBaseConf=[System.__ComObject].InvokeMember("Синоним",[System.Reflection.BindingFlags]::GetProperty,$null,$metadata,$null)
	        $baseConf = ($urlbasesXml.settings.Config | ? {$_.realname -ieq "$realBaseConf"}).name
            $dbInfo.Conf = $baseConf   

            writeLog "Проверка наличия соединений с базой"
            $count = countBaseConnections -comConnector1C $comConnector1C
            if ($count -gt 1) {
                throw "Обнаружено $count подключений к базе. Обновление не будет запущено"
            }
            
            $dbInfo

        } catch {
            if ($error[0] -match "Неправильное имя или пароль пользователя") {
		        $logUser_Password = "\\bit-live.ru\scripts\AutoUpdate\log\logUser_Password.txt"
		        $DisableName = $dbInfo.basename		        
		        Write-Output "$DisableName не правильные логопасы, снимаем с автообновлений" | Out-File -FilePath $logUser_Password -Encoding utf8 -Append
	            $Query = "UPDATE [AutoUpdate].[dbo].[AutoUp] SET Status = 'UpdateDisabled' WHERE BaseName = '$DisableName'"
                sql_Auto -Query $Query | out-null
                throw "$DisableName --> не правильные логопасы, снимаем с автообновлений"
		    }
            throw $error[0]
        }
    }
       
    # Поиск номера следующего обновления. Возвращает null если ничего не найдено либо обновление не требуется.
    function getTrueUpdate ($dbInfo, $updates1CRootPath) {	
        [xml]$updatesXML = Get-Content -Path "$($updates1CRootPath)\$($dbInfo.Conf)\v8cscdsc.xml" -Encoding UTF8
	    [array]$resultUpdatesList = $null
	    $baseVersion = $dbInfo.Version
    
	    if ($baseVersion -match $updatesXML.updateList.LastChild.version -or $baseVersion -match $updatesXML.updateList.LastChild.version.InnerText) {
		    return $null
	    }
	    else {
            foreach ($update in $updatesXML.updateList.update) {
                if ($update.target.Count -eq $null) {continue}
                $update.target | % {
                    if ($_ -eq $baseVersion) {
                        $resultUpdatesList += $update.version;
                    }
                }
		    }
		    if ($resultUpdatesList){            
                $trueUpdate = $resultUpdatesList | Sort-Object -Property InnerText -Descending | select -First 1
			    # Заплатка для нового формата XML
			    if ( $trueUpdate.InnerText ) { return $trueUpdate.InnerText }				
			    else { return $trueUpdate }
		    }
	    }
    }    

    # Получение количества текущих соединений с базой
    function countBaseConnections ($comConnector1C) {
	    if (!($comConnector1C)){ 
		    return		
	    }
        $connections = [System.__ComObject].InvokeMember("ПолучитьСоединенияИнформационнойБазы",[System.Reflection.BindingFlags]::InvokeMethod,$null,$comConnector1C,$null)
	    $connectionsCount = [System.__ComObject].InvokeMember("Количество",[System.Reflection.BindingFlags]::InvokeMethod,$null,$connections,$null)
	    $result = $connectionsCount
        $result
    }

    # Обновление файловой базы
    function updateFileBase($dbInfo, $updates1CRootPath) {
        #throw "FILE update DISABLED"
        $1cv8ExePath = get1CPath
        $BaseUser = $dbInfo.Login
        $BasePassword = $dbInfo.Pass
        $PathToBase = $dbInfo.Path
        $ConfigBase = $dbInfo.Conf
        $BaseVersion = $dbInfo.Version
        $UserBaseAll = $dbInfo.BaseName
        $updateVersion = $dbInfo.NextUpdateVersion
        $Clientid = $dbInfo.ID
        $cfuPath = "$updates1CRootPath\$ConfigBase\$updateVersion\1cv8.cfu"
        if (!(Test-Path "$cfuPath")) {throw "недоступен файл обновления: $cfuPath"}

        [string]$arg1 = 'CONFIG'
        [string]$arg2 = '/F"' + $PathToBase + '"'
        #Update ARGS
		    $arg4 = '/N"' + $BaseUser + '"'
	    $arg5 = '/P"' + $BasePassword + '"'
	
        [string]$arg6 = '/UpdateCfg"' + "$cfuPath" + '"'
        [string]$arg7 = '/UpdateDBCfg'
        [string]$arg8 = '/AU-'
        [string]$arg9 = "'/DisableStartupMessages', '/DisableStartupDialogs'"
	        
        try {        
		    if ($BaseUser -ne "") {
				if ($BasePassword -eq "") {
					#$Args = $arg1, $arg9, $arg2, $arg4, $arg6, $arg7, $arg8
					start-process -FilePath "$1cv8ExePath" -ArgumentList $arg1, $arg9, $arg2, $arg4, $arg6, $arg7, $arg8 -Wait -PassThru -NoNewWindow | Out-Null
				}
				else {
					#$Args = $arg1, $arg9, $arg2, $arg4, $arg5, $arg6, $arg7, $arg8
					start-process -FilePath "$1cv8ExePath" -ArgumentList $arg1, $arg9, $arg2, $arg4, $arg5, $arg6, $arg7, $arg8 -Wait -PassThru -NoNewWindow | Out-Null
				}
			}
			else {
				#$Args = $arg1, $arg9, $arg2, $arg6, $arg7, $arg8
				start-process -FilePath "$1cv8ExePath" -ArgumentList $arg1, $arg9, $arg2, $arg6, $arg7, $arg8 -Wait -PassThru -NoNewWindow | Out-Null
			}
		
		    $message = "Обновление прошло без ошибок"
            #WriteEventlog -Clientid $dbInfo.ID -BaseName $dbInfo.BaseName -ConfigBase $dbInfo.Conf -CurrentVersion $dbInfo.Version -UpdateVersion $dbInfo.NextUpdateVersion -Eventtype Info -EventId 18020 -Message $message                       
            
            $Global:updateResult.currentVersion = "$($dbInfo.Version)"
            $Global:updateResult.updateVersion = "$($dbInfo.NextUpdateVersion)"
            $Global:updateResult.eventType = "Info"
            $Global:updateResult.eventID = "18020"
            $Global:updateResult.message = "$message"

            $dbInfo.Status = $BaseStatus.Done
            $timestamp = Get-Date -Format "dd-MM-yyyy__HH-mm"
            $dbInfo.Date = $timestamp
            $dbInfo.Version = $updateVersion
            $dbInfo.NextUpdateVersion = ""
            writeInfoToDB -dbInfo $dbInfo
            writeLog "Обновление завершено без ошибок"
        } catch {
            $message = $error[0].Exception.Message
            #WriteEventlog -Clientid $dbInfo.ID -BaseName $dbInfo.BaseName -ConfigBase $dbInfo.Conf -CurrentVersion $dbInfo.Version -UpdateVersion $dbInfo.NextUpdateVersion -Eventtype Error -EventId 18010 -Message $message
            
            $Global:updateResult.currentVersion = "$($dbInfo.Version)"
            $Global:updateResult.updateVersion = "$($dbInfo.NextUpdateVersion)"
            $Global:updateResult.eventType = "Error"
            $Global:updateResult.eventID = "18010"
            $Global:updateResult.message = "$message"

            writeLog "$($error[0].Exception.Message)" -error
        }    
    }

    # Обновление SQL базы
    function updateSqlBase($dbInfo, $updates1CRootPath) {
        #throw "SQL update DISABLED"

        $1cv8ExePath = get1CPath
        $serverName = $dbInfo.sqlserver
        $baseName = $dbInfo.BaseName
        $BaseUser = $dbInfo.Login
        $BasePassword = $dbInfo.Pass
        $ConfigBase = $dbInfo.Conf
        $BaseVersion = $dbInfo.Version
        $UserBaseAll = $dbInfo.BaseName
        $updateVersion = $dbInfo.NextUpdateVersion
        $Clientid = $dbInfo.ID
        $cfuPath = "$updates1CRootPath\$ConfigBase\$updateVersion\1cv8.cfu"
        if (!(Test-Path "$cfuPath")) {throw "недоступен файл обновления: $cfuPath"}

        [string]$arg1 = 'DESIGNER'
        [string]$arg2 = '/S"' + $serverName + '\' + $baseName + '"'
		$arg4 = '/N"' + $BaseUser + '"'
	    $arg5 = '/P"' + $BasePassword + '"'	
        [string]$arg6 = '/UpdateCfg"' + "$cfuPath" + '"'
        [string]$arg7 = '/UpdateDBCfg'
        [string]$arg8 = '-server'
        [string]$arg9 = "'/DisableStartupMessages', '/DisableStartupDialogs'"

        try {        
		    if ($BaseUser -ne "") {
				if ($BasePassword -eq "") {
					#$Args = $arg1, $arg9, $arg2, $arg4, $arg6, $arg7, $arg8
					start-process -FilePath "$1cv8ExePath" -ArgumentList $arg1, $arg9, $arg2, $arg4, $arg6, $arg7, $arg8 -Wait -PassThru -NoNewWindow | Out-Null
				}
				else {
					#$Args = $arg1, $arg9, $arg2, $arg4, $arg5, $arg6, $arg7, $arg8
					start-process -FilePath "$1cv8ExePath" -ArgumentList $arg1, $arg9, $arg2, $arg4, $arg5, $arg6, $arg7, $arg8 -Wait -PassThru -NoNewWindow | Out-Null
				}
			}
			else {
				#$Args = $arg1, $arg9, $arg2, $arg6, $arg7, $arg8
				start-process -FilePath "$1cv8ExePath" -ArgumentList $arg1, $arg9, $arg2, $arg6, $arg7, $arg8 -Wait -PassThru -NoNewWindow | Out-Null
			}
		
		    $message = "Обновление прошло без ошибок"
            #WriteEventlog -Clientid $dbInfo.ID -BaseName $dbInfo.BaseName -ConfigBase $dbInfo.Conf -CurrentVersion $dbInfo.Version -UpdateVersion $dbInfo.NextUpdateVersion -Eventtype Info -EventId 18020 -Message $message                       
            
            $Global:updateResult.currentVersion = "$($dbInfo.Version)"
            $Global:updateResult.updateVersion = "$($dbInfo.NextUpdateVersion)"
            $Global:updateResult.eventType = "Info"
            $Global:updateResult.eventID = "18020"
            $Global:updateResult.message = "$message"

            $dbInfo.Status = $BaseStatus.Done
            $timestamp = Get-Date -Format "dd-MM-yyyy__HH-mm"
            $dbInfo.Date = $timestamp
            $dbInfo.Version = $updateVersion
            $dbInfo.NextUpdateVersion = ""
            writeInfoToDB -dbInfo $dbInfo
            writeLog "Обновление завершено без ошибок"
        } catch {
            $message = $error[0].Exception.Message
            #WriteEventlog -Clientid $dbInfo.ID -BaseName $dbInfo.BaseName -ConfigBase $dbInfo.Conf -CurrentVersion $dbInfo.Version -UpdateVersion $dbInfo.NextUpdateVersion -Eventtype Error -EventId 18010 -Message $message
            $Global:updateResult.currentVersion = "$($dbInfo.Version)"
            $Global:updateResult.updateVersion = "$($dbInfo.NextUpdateVersion)"
            $Global:updateResult.eventType = "Error"
            $Global:updateResult.eventID = "18010"
            $Global:updateResult.message = "$message"
            writeLog "$($error[0].Exception.Message)" -error
        } 
    }

    function validateInput {
        if ($dbInfo.Path -eq $null -and $dbInfo.sqlserver -eq $null) {
            throw "Данные в dbInfo неверны или отсутствуют"
        }
        
        if ($dbInfo.Path -eq "" -and $dbInfo.sqlserver -eq "") {
            throw "Данные в dbInfo неверны или отсутствуют"
        }

        if (!(Test-Path "$($dbModulePath)")) {
            throw "неверный путь в параметре dbModulePath: [$dbModulePath]"
        }
        if (!(Test-Path "$($updates1CRootPath)")) {
            throw "неверный путь в параметре updates1CRootPath: [$updates1CRootPath]"
        }
        if (!(Test-Path "$($urlbasesXmlPath)")) {
            throw "неверный путь в параметре urlbasesXmlPath: [$urlbasesXmlPath]"
        }        
        checkComAppForUpdate   
        get1CPath | Out-Null
    }

    # Поиск пути к 1cv8.exe с использованием COM-коннектора 
    function get1CPath() {
        [string]$CLSid83 = 'HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{181E893D-73A4-4722-B61D-D604B3D67D47}\InprocServer32'
        $Path1cExe = ((Get-Item -path $CLSid83).GetValue('')) -replace "comcntr.dll","1cv8.exe"
        if (!(Test-Path $Path1cExe)) {throw "Не удалось найти путь к 1сv8.exe"}
        $Path1cExe
    }

    # Записывает указанную информацию о базе 1С в БД
    function writeInfoToDB ($dbInfo) {    
        $Query = "UPDATE [AutoUpdate].[dbo].[AutoUp] SET Conf = `'$($dbInfo.Conf)`', Status = `'$($dbInfo.Status)`', Date = `'$($dbInfo.Date)`', Version = `'$($dbInfo.Version)`', NextUpdateVersion = `'$($dbInfo.NextUpdateVersion)`' WHERE BaseName = `'$($dbInfo.BaseName)`'"
        sql_Auto -Query $Query | out-null
    }

    # Проверка существования отдельного COM приложения для обновлений и его создания
    function checkComAppForUpdate ([switch]$rerun)  {
        try {
            $comDllRootPath = "C:\Program Files (x86)\1cv8\"
            if (!(Test-Path "$comDllRootPath")) {throw "Нет установленных 1cv8"}   

            $salt = generateSalt
            $defaultConnectorAppName = "V83COMConnector"
            $connectorAppName = "UPDATE.V83.COMConnector.$salt"
            writeLog "CREATE_COM --> Поиск COM+ приложения [$connectorAppName]"
            $comAdmin = New-Object -comobject COMAdmin.COMAdminCatalog
            $apps = $comAdmin.GetCollection("Applications")
            $apps.Populate()
            if (($apps | select -ExpandProperty name) -contains $connectorAppName) {
                writeLog "CREATE_COM --> COM+ приложение [$connectorAppName] найдено"
                return
            }
            writeLog "CREATE_COM --> COM+ приложение [$connectorAppName] не обнаружено. Выполняется создание"

            writeLog "CREATE_COM --> Поиск существующих коннекторов 1С"
            $isFounded = $false
            foreach ($app in $apps) {
                $components = $apps.GetCollection("Components",$app.Key)
                $components.Populate()
                foreach ($comp in $components) {
                    if ($comp.Value("DLL") -match ".*comcntr.dll$") {
                        $defaultApp = $app
                        $isFounded = $true
                        writeLog "CREATE_COM --> Приложение найдено: [$($app.Name)]"
                        break;
                    }
                }
                if ($isFounded) {break;}
            }

            if (!$isFounded) {
                writeLog "CREATE_COM --> Не найдено стандартное COM+ приложение 1C. Создание"
                $newApp = $apps.Add()
                $newApp.Value("Name") = $defaultConnectorAppName
			    $apps.SaveChanges() | Out-Null

                $comDll = (Get-ChildItem -Path $comDllRootPath -Filter comcntr.dll -Recurse | Select-Object FullName -Last 1).FullName
			    $comAdmin.InstallComponent($defaultConnectorAppName, $comDll, "", "")
                cmd /c (regsvr32 "$comDll" /i /s) 
                $defaultApp = $newApp
                writeLog "CREATE_COM --> COM-коннектор [$defaultConnectorAppName] создан с удостоверением [$($defaultApp.Value("Identity"))]"
            }

            $newApp = $apps.Add()
		    $newApp.Value("Name") = $connectorAppName
		    $newApp.Value("ApplicationAccessChecksEnabled") = 0
		    $newApp.Value("AccessChecksLevel") = 1
		    $newApp.Value("SRPEnabled") = 1
		    $newApp.Value("SRPTrustLevel") = 262144				
		    $newApp.Value("Identity") = $($creds.UserName)
		    $newApp.Value("Password") = [System.Net.NetworkCredential]::new("", $creds.Password).Password
		    $apps.SaveChanges() | Out-Null
            $Global:connectorName = $connectorAppName
        
            $components = $apps.GetCollection("Components",$defaultApp.Key)
            $components.Populate()

            $sourceAppId = $defaultApp.Key
            $sourceCompId = ($components | select -First 1).key
            $destAppId = $newApp.Key
            $newProgId = $connectorAppName
            $newClsid = "" #"{C7C8B9AE-751E-4AAD-B392-A3939129C507}"
            $comAdmin.AliasComponent($sourceAppId, $sourceCompId, $destAppId, $newProgId, $newClsid)
            writeLog "CREATE_COM --> COM+ приложение [$connectorAppName] создано с удостоверением [$($creds.UserName)]"
        }
        catch {
            if (!$rerun -and (
                    $($error[0].Exception.Message) -match "Исключение из HRESULT: 0x80110401" -or
                    $($error[0].Exception.Message) -match "Исключение из HRESULT: 0x8011045C"
                )) {
                writeLog "CREATE_COM --> При создании коннектора произошла ошибка [$($error[0].Exception.Message)]" -error
                writeLog "CREATE_COM --> Вероятно коннектор [$($defaultApp.name)] поврежден. Выполняется пересоздание" -error
                removeComConnector -connectorAppName $connectorAppName
                removeComConnector -connectorAppName $($defaultApp.Name)
                removeComConnector -connectorAppName $defaultConnectorAppName
                checkComAppForUpdate -rerun

            } else {
                throw $($error[0])
            }               
        }                
    }

    function removeComConnector ($connectorAppName) {
        if ($connectorAppName -eq $null -or $connectorAppName -eq "" -or $connectorAppName.Length -lt 6) {return}
        writeLog "DELETE_COM --> Удаление COM-коннектора [$connectorAppName]"
        $comAdmin = New-Object -comobject COMAdmin.COMAdminCatalog
        $apps = $comAdmin.GetCollection("Applications")
        $apps.Populate()

        for ($i = 0; $i -lt $apps.Count; $i++) {
            if ($apps.item($i).name -eq "$connectorAppName") {
                $apps.Remove($i)
                $apps.SaveChanges() | Out-Null
                writeLog "DELETE_COM --> COM-коннектор [$connectorAppName] удалён"
                return
            }            
        }
        
        writeLog "DELETE_COM --> COM-коннектор [$connectorAppName] не найден"
    }

    function generateSalt {
        function Get-RandomChars {
	        param(        
		        $length,
		        $characters
	        )
	        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
	        $private:ofs=""
	        [String]$characters[$random]
        }

	    #$salt = Get-RandomChars -length 6 -characters 'abcdefghiklmnprstuvwxyz'
	    $salt = Get-RandomChars -length 5 -characters '1234567890'	
        $len = $salt.length -1
	    $indizes = Get-Random -InputObject (0..$len) -Count $len	
	    $private:ofs=''
	    return [String]$salt[$indizes]
    }


#endregion

#region Main

try {
    $timeStart = Get-Date    
    Import-Module "$dbModulePath"

    writeLog "Проверка параметров"
    validateInput

    $Global:updateResult = New-Object -TypeName PSObject -Property @{
                clientID = "$($dbInfo.ID)"
                baseName = "$($dbInfo.BaseName)"
                configBase = "$($dbInfo.Conf)"
                currentVersion = "$($dbInfo.Version)"
                updateVersion = "$($dbInfo.NextUpdateVersion)"
                eventType = "Info"
                eventID = "18020"
                message = ""
            }

    if ($dbInfo.BaseType -eq "File") {
        writeLog "Получение информации о файловой базе"
        $dbInfo = getFileBaseInformation -dbInfo $dbInfo -urlbasesXmlPath $urlbasesXmlPath      
    } elseif ($dbInfo.BaseType -eq "SQL") {
        writeLog "Получение информации о SQL базе"
        $dbInfo = getSqlBaseInformation -dbInfo $dbInfo -urlbasesXmlPath $urlbasesXmlPath
    } else {throw "Не удалось определить тип БД. dbInfo.BaseType=[$($dbInfo.BaseType)]"}

    $dbInfo = checkIsUpdateRequired -dbInfo $dbInfo  
    $dbInfo.Status = $BaseStatus.Updating
    writeInfoToDB -dbInfo $dbInfo
    if ($dbInfo.BaseType -eq "File") {
        writeLog "Запускается обновление файловой базы"
        updateFileBase -dbInfo $dbInfo -updates1CRootPath "$updates1CRootPath"
    } else {
        writeLog "Запускается обновление SQL базы"
        updateSqlBase -dbInfo $dbInfo -updates1CRootPath $updates1CRootPath
    }
    removeComConnector -connectorAppName $Global:connectorName
    $timeEnd = Get-Date
    $totalTime = $timeEnd - $timeStart
    writeLog "Время выполнения: $($totalTime.Minutes)m. $($totalTime.Seconds)s."
    $Global:updateResult
} 
catch {
    writeLog "$($error[0].Exception.Message)" -error
    removeComConnector -connectorAppName $Global:connectorName
    $Global:updateResult.message = "$($error[0].Exception.Message)"
    $Global:updateResult
}

#endregion




<#test object

$dbinfo = New-Object -TypeName pscustomobject -Property @{id="";basename="";path="";conf="";status="";date="";login="";version="";nextupdateversion=""}

$dbInfo.ID = "0000"
$dbInfo.BaseName = "id0000_test"
$dbInfo.Path = "\\tstrdsh01\testDB\InfoBase4"
$dbInfo.Conf = "Accounting3"
$dbInfo.Status = "WaitingForUpdate"
$dbInfo.Date = "14-10-2020__12-07"
$dbInfo.Login = "Администратор"
$dbInfo.Version = ""
$dbInfo.NextUpdateVersion = ""

#>
