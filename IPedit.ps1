
#変数設定
$servername = "●●●●"
$InputFile = "▲▲▲▲"
$domain = ".■■■■"
$pclist = Import-CSV -Path $InputFile  


$global:exist_host = @()
$global:not_exist_host =@()

$machine_delete_pclist_exist = @()
$global:remove_check_flag = 0
$global:add_check_flag = 0
$global:csvinfo_check_flag = 0
$global:all_macaddresslist = 0

$global:change_count = 0
$global:csvdata_check = 0

$reference = 0
$reference_do = 0
$reference_pclist = @()
$machine_delete = 0
$machine_delete_do = 0
$machine_delete_pclist = @()
$add = 0
$add_do = 0
$add_pclist = @()
$IPdelete = 0
$global:IPdelete_do = 0
$IPdelete_pclist = @()


#選択肢の作成
$typename = "System.Management.Automation.Host.ChoiceDescription"
$yes = new-object $typename("&Yes","実行する")
$no  = new-object $typename("&No","実行しない")

#選択肢コレクションの作成
$assembly= $yes.getType().AssemblyQualifiedName
$choice = new-object "System.Collections.ObjectModel.Collection``1[[$assembly]]"
$choice.add($yes)
$choice.add($no)


#現在の予約IPリストを取得
function GetAllReserveIPList($reserve_iplist){
    try{

            $reserve_iplist = Get-DhcpServerv4Scope -computername $servername | Get-DhcpServerv4Reservation -computername $servername
            return $reserve_iplist
        
    }catch{
        Write-Output "起動権限をご確認ください"
        Write-Output `n
        pause
        exit
    }
}



#対象端末の予約状況取得
function GetPCReserveList($get_pclist,$get_iplist){
    
    foreach ($pc in $get_pclist){
        try{
            $checkname = $pc.name.tolower()
            $iplist_hostname = $get_iplist.name -replace ($domain ,"")
            if($iplist_hostname -match "^$checkname"){
                
                $global:exist_host += $get_iplist  | Where-Object { $_.name -match $checkname}

            }else{
                $global:not_exist_host += $pc

            }
            

        }catch{
            Write-Output $PSItem.Exception.Message
            Write-Output "予約リストと対象端末の照会に失敗"
            Write-Output `n
            pause
        }
    }

    if($exist_host -ne $null){
        Write-Output $exist_host
        Write-Output `n

    }
    if($not_exist_host -ne $null){
        $temp_text = $pc.name + ":" + $pc.ipaddress + "は予約IP登録がありません"
        Write-Output $temp_text
        Write-Output `n

    }
           
    if($not_exist_host.count -eq $get_pclist.count){
        
        Write-Output "対象端末をサーバ上で確認できませんでした"

        Write-Output `n
        $global:csvinfo_check_flag = 1
        pause

    }
    
}

function GetPCReserveList_IP($get_pclist,$get_iplist){
    
    foreach ($pc in $get_pclist){
        try{
            $chekname = $pc.ipaddress
            if($get_iplist.ipaddress -match "^$checkname"){
                
                $global:exist_host += $get_iplist  | Where-Object { $_.ipaddress -match $pc.ipaddress}

            }else{
                $global:not_exist_host += $pc

            }
            

        }catch{
            Write-Output $PSItem.Exception.Message
            Write-Output "予約リストと対象端末の照会に失敗"
            Write-Output `n
            pause
        }
    }

    if($exist_host -ne $null){
        Write-Output $exist_host
        Write-Output `n

    }
    if($not_exist_host -ne $null){
        $temp_text = $pc.name + ":" + $pc.ipaddress + "は予約IP登録がありません"
        Write-Output $temp_text
        Write-Output `n

    }
           
    if($not_exist_host.count -eq $get_pclist.count){
        
        Write-Output "対象端末をサーバ上で確認できませんでした"

        Write-Output `n
        $global:csvinfo_check_flag = 1
        pause

    }
    
}



#予約IPの削除
function RemoveReserveIP{
    Write-Output "◆予約IPを削除します"

    try{
            foreach($eh in $exist_host){
                $eh | Remove-DhcpServerv4Reservation -computername $servername -WhatIf
                $eh | Remove-DhcpServerv4Reservation -computername $servername
            }
            $global:IPlist_fix_check_flag =1
      
    }catch{
        Write-Output "予約IP削除に失敗しました"
        Write-Output `n
    }
}

#各スコープのポリシーから対象端末MACアドレスを削除する
function RemovePolicyMacAddress{
    $global:change_count = 0

    for ( $i = 0; $i -lt $exist_host.length;$i++){
        $global:all_macaddresslist = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}} #対象端末のMACアドレス取得


        $temp_macaddresslist = New-Object System.Collections.ArrayList
        $temp_macaddresslist.addrange($all_macaddresslist[$i].macaddress.split(";"))

        if($temp_macaddresslist -contains $exist_host[$i].clientid.tolower()){
            $temp_count = $temp_macaddresslist.Count - 1
            $temp_text = "現在の登録数：" + $temp_count + " 端末名:" + $exist_host[$i].name + "     スコープ:" + $exist_host[$i].scopeid + "     MACアドレス:" + $exist_host[$i].clientid.tolower() + " の登録を確認できました"
            Write-Output $temp_text

            $temp_macaddresslist.remove($exist_host[$i].clientid.tolower()) | Out-Null

            [string[]]$macadd=@()
            $macadd = $temp_macaddresslist
            
            try{
                set-DhcpServerv4Policy -computername $servername -ScopeId $exist_host[$i].scopeid -Name Allow  -Condition AND -MacAddress $macadd -WhatIf
                set-DhcpServerv4Policy -computername $servername -ScopeId $exist_host[$i].scopeid -Name Allow  -Condition AND -MacAddress $macadd
                $global:remove_check_flag = 1
                $global:change_count += 1
            }catch{
                $temp_text = "ポリシーリストからMacアドレス削除に失敗しました。" +  $temp_macaddresslist.Count + " 端末名:" + $exist_host[$i].name + "     スコープ:" + $exist_host[$i].scopeid + "     MACアドレス:" + $exist_host[$i].clientid.tolower() + " は削除されませんでした。"
                Write-Output $temp_text
                Write-Output $PSItem.Exception.Message
                pause
            }
        }else{
            $temp_text = "端末名:" + $exist_host[$i].name + "     スコープ:" + $exist_host[$i].scopeid + "     MACアドレス:" + $exist_host[$i].clientid + " の登録がありません"
            Write-Output $temp_text
            $global:check_flag = 0
        }   
    }
}

function AddPolicyMacAddress{
    $global:change_count = 0

    for ( $i = 0; $i -lt $exist_host.length;$i++){
        $global:all_macaddresslist = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}} #対象端末のMACアドレス取得

        $temp_macaddresslist = New-Object System.Collections.ArrayList
        $temp_macaddresslist.addrange($all_macaddresslist[$i].macaddress.split(";"))

        if($temp_macaddresslist -notcontains $exist_host[$i].clientid.tolower()){
            $temp_count = $temp_macaddresslist.Count - 1
            $temp_text = "現在の登録数：" + $temp_count + " 端末名:" + $exist_host[$i].name + "     スコープ:" + $exist_host[$i].scopeid + "     MACアドレス:" + $exist_host[$i].clientid.tolower() + " の登録がないことを確認できました"
            Write-Output $temp_text

            $temp_macaddresslist.add($exist_host[$i].clientid.tolower()) | Out-Null
   
            [string[]]$macadd=@()
            $macadd = $temp_macaddresslist
            
              
            try{
            
                set-DhcpServerv4Policy -computername $servername -ScopeId $exist_host[$i].scopeid -Name Allow  -Condition AND -MacAddress $macadd -WhatIf
                set-DhcpServerv4Policy -computername $servername -ScopeId $exist_host[$i].scopeid -Name Allow  -Condition AND -MacAddress $macadd
                $global:add_check_flag = 1
                $global:change_count += 1
            }catch{
                $temp_text = "ポリシーリストにMacアドレス追加に失敗しました。"
                Write-Output $temp_text
                Write-Output $PSItem.Exception.Message
                pause
            }


            
            
            }else{
                $temp_text = "端末名:" + $exist_host[$i].name + "     スコープ:" + $exist_host[$i].scopeid + "     MACアドレス:" + $exist_host[$i].clientid + " の登録があります"
                Write-Output $temp_text
                $global:check_flag = 0
            }  

           




    }
     

}


#ポリシーのMacアドレス数の確認
function CheckPolicyMacNo{
    if($remove_check_flag -ne 0){
        Write-Output "◆ポリシー登録状況確認"
        $all_macaddresslist_fix = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}}
 
        for ( $j = 0; $j -lt $exist_host.Length;$j++){
        
            if($all_macaddresslist[$j].macaddress.count - $change_count -eq $all_macaddresslist_fix[$j].macaddress.count){
            
                $temp_text = "登録数：" + ($all_macaddresslist_fix[$j].macaddress.count - 1) +  " スコープ:" + $exist_host[$j].scopeid + "の削除は正常に機能しました"
                Write-Output $temp_text

            }else{
                 $temp_text = "登録数：" + ($all_macaddresslist_fix[$j].macaddress.count - 1) + " スコープ:" + $exist_host[$j].scopeid + "で不整合が起きています"
                Write-Output $temp_text

            }
         }
    }

       if($add_check_flag -ne 0){
        Write-Output "◆ポリシー登録状況確認"
        $all_macaddresslist_fix = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}}
 
        for ( $j = 0; $j -lt $exist_host.Length;$j++){
        
            if($all_macaddresslist[$j].macaddress.count + $change_count -eq $all_macaddresslist_fix[$j].macaddress.count){
            
                $temp_text = "登録数：" + ($all_macaddresslist_fix[$j].macaddress.count - 1) +  " スコープ:" + $exist_host[$j].scopeid + "の追加は正常に機能しました"
                Write-Output $temp_text

            }else{
                 $temp_text = "登録数：" + ($all_macaddresslist_fix[$j].macaddress.count - 1) + " スコープ:" + $exist_host[$j].scopeid + "で不整合が起きています"
                Write-Output $temp_text

            }
         }
    }
}



##CSVのインポートチェック
 function CsvImportCheck($checklist){
        if($csvinfo_check_flag -ne 0){
            Write-Output "csvのMACアドレス、IPアドレスを利用して確認します"

            
            $global:exist_host = $checklist
            Write-Output $exist_host | Format-table -Property action,name,clientid,ipaddress,scopeid


            foreach($n in $exist_host){
                if($IPdelete_do -eq 0){
                    if([string]::IsNullOrEmpty($n.clientid)){
                        $temp_text = $n.name + "のMACアドレスがありません" 
                        Write-Output $temp_text
                        $global:csvdata_check = 1
                    }
                
                }
                if([string]::IsNullOrEmpty($n.ipaddress)){
                    $temp_text = $n.name + "のIPアドレスがありません" 
                    Write-Output $temp_text
                    $global:csvdata_check = 1
                
                }
                if([string]::IsNullOrEmpty($n.scopeid)){
                    if($IPdelete_do -eq 0){
                        $temp_text = $n.name + "のスコープIDがありません" 
                        Write-Output $temp_text
                        $global:csvdata_check = 1
                    }
                
                }
            }
            

        }

        if($csvdata_check -ne 0){
            Write-Output "csvデータに不整合があります。見直してください"
            pause
            Write-Output "処理を中止します"
            exit
        }
    }



####実行####

##データ取得##
if($iplist -eq $null -Or $IPlist_fix_check_flag -ne 0){
    $iplist = GetAllReserveIPList $iplist
}
$global:IPlist_fix_check_flag =0
##処理数確認開始##

for($k=0; $k -lt $pclist.action.Length;$k++){
    if($pclist[$k].action -eq 1){
        $reference += 1
        $reference_pclist += $pclist[$k]

    }
     if($pclist[$k].action -eq 2){
        $machine_delete += 1
        $machine_delete_pclist += $pclist[$k]

    }
    if($pclist[$k].action -eq 3){
        $add += 1
        $add_pclist += $pclist[$k]
    }

    if($pclist[$k].action -eq 4){
        $IPdelete += 1
        $IPdelete_pclist += $pclist[$k]
    }
 
}


$temp_text = "●全処理数:" + $pclist.action.length
Write-Output $temp_text


if($reference -ne 0){
    $temp_text = "●参照数:" + $reference
    Write-Output $temp_text
    $reference_do = 1

}
if($machine_delete -ne 0){
    $temp_text = "●機器削除数:" + $machine_delete
    Write-Output $temp_text
    $machine_delete_do = 1
}
if($add -ne 0){
    $temp_text = "●追加数:" + $add
    Write-Output $temp_text
    $add_do = 1
}
if($IPdelete -ne 0){
    $temp_text = "●IP削除数:" + $IPdelete
    Write-Output $temp_text
    $global:IPdelete_do = 1
}

Write-Output `n


if($reference_do -eq 1){
    Write-Output "◆IPの参照処理を始めます"
    GetPCReserveList  $reference_pclist $iplist
}



if($machine_delete_do -eq 1){
    ##削除処理開始##
    Write-Output "◆機器の削除処理を始めます"
    GetPCReserveList $machine_delete_pclist $iplist #対象端末の予約状況
    CsvImportCheck $machine_delete_pclist
   
    #選択プロンプトの表示
    $answer = $host.ui.PromptForChoice("<実行確認>","実行しますか？",$choice,0)

    #$execute = Read-Host "削除を実行しますか? y/n"
    Write-Output `n

    if ($answer -eq "0"){
        if($csvinfo_check_flag -eq 0){
            RemoveReserveIP #予約IPの削除
        }
        Write-Output `n
        Write-Output "◆ポリシーからMACアドレスを削除します"
        RemovePolicyMacAddress #各スコープのポリシーから対象端末MACアドレスを削除する
        Write-Output `n
        Write-Output "◆処理終了しました"
        Write-Output `n
    }else{
        Write-Output "削除がキャンセルされました"
    }
    Write-Output `n
   CheckPolicyMacNo
    Write-Output `n
}




if($add_do -eq 1){
    ##追加処理開始##
    Write-Output "◆IPの追加処理を始めます"
    $global:csvinfo_check_flag = 1
    CsvImportCheck $add_pclist


    #選択プロンプトの表示
    $answer = $host.ui.PromptForChoice("<実行確認>","実行しますか？",$choice,0)


    if ($answer -eq "0"){
        

            foreach($eh in $exist_host){

                if($eh.ipaddress -ne  $iplist ){
                    try{
                        $eh | Add-DhcpServerv4Reservation -computername $servername -type dhcp -WhatIf
                        $eh | Add-DhcpServerv4Reservation -computername $servername -type dhcp 
                        $global:IPlist_fix_check_flag =1
                    }catch{
                        $temp_text = "予約IPアドレス追加に失敗しました。"
                        Write-Output $temp_text
                        pause
                    }
                }else{
                    $temp_text = $eh.ipaddress + "はすでに登録があります"
                    write-output $temp_text
                    pause
                }
                }

            
        

        Write-Output `n
        Write-Output "◆ポリシーにMACアドレスを追加します"
        AddPolicyMacAddress #各スコープのポリシーから対象端末MACアドレスを削除する
        Write-Output `n
        Write-Output "◆処理終了しました"
        Write-Output `n
    }else{
        Write-Output "追加がキャンセルされました"
    }
    Write-Output `n
   CheckPolicyMacNo
    Write-Output `n
}

if($IPdelete_do -eq 1){
    ##削除処理開始##
    Write-Output "◆IPの削除処理を始めます"
    GetPCReserveList_IP $IPdelete_pclist $iplist
   
    $global:all_macaddresslist = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}} #対象端末のMACアドレス取得

    #選択プロンプトの表示
    $answer = $host.ui.PromptForChoice("<実行確認>","実行しますか？",$choice,0)

    #$execute = Read-Host "削除を実行しますか? y/n"
    Write-Output `n

    if ($answer -eq "0"){
        if($csvinfo_check_flag -eq 0){
            RemoveReserveIP #予約IPの削除
        }
        Write-Output `n
        Write-Output "◆ポリシーからMACアドレスを削除します"
        RemovePolicyMacAddress #各スコープのポリシーから対象端末MACアドレスを削除する
        Write-Output `n
        Write-Output "◆処理終了しました"
        Write-Output `n
    }else{
        Write-Output "削除がキャンセルされました"
    }
    Write-Output `n
   CheckPolicyMacNo
    Write-Output `n
}

pause




