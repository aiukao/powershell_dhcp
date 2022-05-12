
#�ϐ��ݒ�
$servername = "��������"
$InputFile = "��������"
$domain = ".��������"
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


#�I�����̍쐬
$typename = "System.Management.Automation.Host.ChoiceDescription"
$yes = new-object $typename("&Yes","���s����")
$no  = new-object $typename("&No","���s���Ȃ�")

#�I�����R���N�V�����̍쐬
$assembly= $yes.getType().AssemblyQualifiedName
$choice = new-object "System.Collections.ObjectModel.Collection``1[[$assembly]]"
$choice.add($yes)
$choice.add($no)


#���݂̗\��IP���X�g���擾
function GetAllReserveIPList($reserve_iplist){
    try{

            $reserve_iplist = Get-DhcpServerv4Scope -computername $servername | Get-DhcpServerv4Reservation -computername $servername
            return $reserve_iplist
        
    }catch{
        Write-Output "�N�����������m�F��������"
        Write-Output `n
        pause
        exit
    }
}



#�Ώے[���̗\��󋵎擾
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
            Write-Output "�\�񃊃X�g�ƑΏے[���̏Ɖ�Ɏ��s"
            Write-Output `n
            pause
        }
    }

    if($exist_host -ne $null){
        Write-Output $exist_host
        Write-Output `n

    }
    if($not_exist_host -ne $null){
        $temp_text = $pc.name + ":" + $pc.ipaddress + "�͗\��IP�o�^������܂���"
        Write-Output $temp_text
        Write-Output `n

    }
           
    if($not_exist_host.count -eq $get_pclist.count){
        
        Write-Output "�Ώے[�����T�[�o��Ŋm�F�ł��܂���ł���"

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
            Write-Output "�\�񃊃X�g�ƑΏے[���̏Ɖ�Ɏ��s"
            Write-Output `n
            pause
        }
    }

    if($exist_host -ne $null){
        Write-Output $exist_host
        Write-Output `n

    }
    if($not_exist_host -ne $null){
        $temp_text = $pc.name + ":" + $pc.ipaddress + "�͗\��IP�o�^������܂���"
        Write-Output $temp_text
        Write-Output `n

    }
           
    if($not_exist_host.count -eq $get_pclist.count){
        
        Write-Output "�Ώے[�����T�[�o��Ŋm�F�ł��܂���ł���"

        Write-Output `n
        $global:csvinfo_check_flag = 1
        pause

    }
    
}



#�\��IP�̍폜
function RemoveReserveIP{
    Write-Output "���\��IP���폜���܂�"

    try{
            foreach($eh in $exist_host){
                $eh | Remove-DhcpServerv4Reservation -computername $servername -WhatIf
                $eh | Remove-DhcpServerv4Reservation -computername $servername
            }
            $global:IPlist_fix_check_flag =1
      
    }catch{
        Write-Output "�\��IP�폜�Ɏ��s���܂���"
        Write-Output `n
    }
}

#�e�X�R�[�v�̃|���V�[����Ώے[��MAC�A�h���X���폜����
function RemovePolicyMacAddress{
    $global:change_count = 0

    for ( $i = 0; $i -lt $exist_host.length;$i++){
        $global:all_macaddresslist = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}} #�Ώے[����MAC�A�h���X�擾


        $temp_macaddresslist = New-Object System.Collections.ArrayList
        $temp_macaddresslist.addrange($all_macaddresslist[$i].macaddress.split(";"))

        if($temp_macaddresslist -contains $exist_host[$i].clientid.tolower()){
            $temp_count = $temp_macaddresslist.Count - 1
            $temp_text = "���݂̓o�^���F" + $temp_count + " �[����:" + $exist_host[$i].name + "     �X�R�[�v:" + $exist_host[$i].scopeid + "     MAC�A�h���X:" + $exist_host[$i].clientid.tolower() + " �̓o�^���m�F�ł��܂���"
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
                $temp_text = "�|���V�[���X�g����Mac�A�h���X�폜�Ɏ��s���܂����B" +  $temp_macaddresslist.Count + " �[����:" + $exist_host[$i].name + "     �X�R�[�v:" + $exist_host[$i].scopeid + "     MAC�A�h���X:" + $exist_host[$i].clientid.tolower() + " �͍폜����܂���ł����B"
                Write-Output $temp_text
                Write-Output $PSItem.Exception.Message
                pause
            }
        }else{
            $temp_text = "�[����:" + $exist_host[$i].name + "     �X�R�[�v:" + $exist_host[$i].scopeid + "     MAC�A�h���X:" + $exist_host[$i].clientid + " �̓o�^������܂���"
            Write-Output $temp_text
            $global:check_flag = 0
        }   
    }
}

function AddPolicyMacAddress{
    $global:change_count = 0

    for ( $i = 0; $i -lt $exist_host.length;$i++){
        $global:all_macaddresslist = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}} #�Ώے[����MAC�A�h���X�擾

        $temp_macaddresslist = New-Object System.Collections.ArrayList
        $temp_macaddresslist.addrange($all_macaddresslist[$i].macaddress.split(";"))

        if($temp_macaddresslist -notcontains $exist_host[$i].clientid.tolower()){
            $temp_count = $temp_macaddresslist.Count - 1
            $temp_text = "���݂̓o�^���F" + $temp_count + " �[����:" + $exist_host[$i].name + "     �X�R�[�v:" + $exist_host[$i].scopeid + "     MAC�A�h���X:" + $exist_host[$i].clientid.tolower() + " �̓o�^���Ȃ����Ƃ��m�F�ł��܂���"
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
                $temp_text = "�|���V�[���X�g��Mac�A�h���X�ǉ��Ɏ��s���܂����B"
                Write-Output $temp_text
                Write-Output $PSItem.Exception.Message
                pause
            }


            
            
            }else{
                $temp_text = "�[����:" + $exist_host[$i].name + "     �X�R�[�v:" + $exist_host[$i].scopeid + "     MAC�A�h���X:" + $exist_host[$i].clientid + " �̓o�^������܂�"
                Write-Output $temp_text
                $global:check_flag = 0
            }  

           




    }
     

}


#�|���V�[��Mac�A�h���X���̊m�F
function CheckPolicyMacNo{
    if($remove_check_flag -ne 0){
        Write-Output "���|���V�[�o�^�󋵊m�F"
        $all_macaddresslist_fix = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}}
 
        for ( $j = 0; $j -lt $exist_host.Length;$j++){
        
            if($all_macaddresslist[$j].macaddress.count - $change_count -eq $all_macaddresslist_fix[$j].macaddress.count){
            
                $temp_text = "�o�^���F" + ($all_macaddresslist_fix[$j].macaddress.count - 1) +  " �X�R�[�v:" + $exist_host[$j].scopeid + "�̍폜�͐���ɋ@�\���܂���"
                Write-Output $temp_text

            }else{
                 $temp_text = "�o�^���F" + ($all_macaddresslist_fix[$j].macaddress.count - 1) + " �X�R�[�v:" + $exist_host[$j].scopeid + "�ŕs�������N���Ă��܂�"
                Write-Output $temp_text

            }
         }
    }

       if($add_check_flag -ne 0){
        Write-Output "���|���V�[�o�^�󋵊m�F"
        $all_macaddresslist_fix = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}}
 
        for ( $j = 0; $j -lt $exist_host.Length;$j++){
        
            if($all_macaddresslist[$j].macaddress.count + $change_count -eq $all_macaddresslist_fix[$j].macaddress.count){
            
                $temp_text = "�o�^���F" + ($all_macaddresslist_fix[$j].macaddress.count - 1) +  " �X�R�[�v:" + $exist_host[$j].scopeid + "�̒ǉ��͐���ɋ@�\���܂���"
                Write-Output $temp_text

            }else{
                 $temp_text = "�o�^���F" + ($all_macaddresslist_fix[$j].macaddress.count - 1) + " �X�R�[�v:" + $exist_host[$j].scopeid + "�ŕs�������N���Ă��܂�"
                Write-Output $temp_text

            }
         }
    }
}



##CSV�̃C���|�[�g�`�F�b�N
 function CsvImportCheck($checklist){
        if($csvinfo_check_flag -ne 0){
            Write-Output "csv��MAC�A�h���X�AIP�A�h���X�𗘗p���Ċm�F���܂�"

            
            $global:exist_host = $checklist
            Write-Output $exist_host | Format-table -Property action,name,clientid,ipaddress,scopeid


            foreach($n in $exist_host){
                if($IPdelete_do -eq 0){
                    if([string]::IsNullOrEmpty($n.clientid)){
                        $temp_text = $n.name + "��MAC�A�h���X������܂���" 
                        Write-Output $temp_text
                        $global:csvdata_check = 1
                    }
                
                }
                if([string]::IsNullOrEmpty($n.ipaddress)){
                    $temp_text = $n.name + "��IP�A�h���X������܂���" 
                    Write-Output $temp_text
                    $global:csvdata_check = 1
                
                }
                if([string]::IsNullOrEmpty($n.scopeid)){
                    if($IPdelete_do -eq 0){
                        $temp_text = $n.name + "�̃X�R�[�vID������܂���" 
                        Write-Output $temp_text
                        $global:csvdata_check = 1
                    }
                
                }
            }
            

        }

        if($csvdata_check -ne 0){
            Write-Output "csv�f�[�^�ɕs����������܂��B�������Ă�������"
            pause
            Write-Output "�����𒆎~���܂�"
            exit
        }
    }



####���s####

##�f�[�^�擾##
if($iplist -eq $null -Or $IPlist_fix_check_flag -ne 0){
    $iplist = GetAllReserveIPList $iplist
}
$global:IPlist_fix_check_flag =0
##�������m�F�J�n##

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


$temp_text = "���S������:" + $pclist.action.length
Write-Output $temp_text


if($reference -ne 0){
    $temp_text = "���Q�Ɛ�:" + $reference
    Write-Output $temp_text
    $reference_do = 1

}
if($machine_delete -ne 0){
    $temp_text = "���@��폜��:" + $machine_delete
    Write-Output $temp_text
    $machine_delete_do = 1
}
if($add -ne 0){
    $temp_text = "���ǉ���:" + $add
    Write-Output $temp_text
    $add_do = 1
}
if($IPdelete -ne 0){
    $temp_text = "��IP�폜��:" + $IPdelete
    Write-Output $temp_text
    $global:IPdelete_do = 1
}

Write-Output `n


if($reference_do -eq 1){
    Write-Output "��IP�̎Q�Ə������n�߂܂�"
    GetPCReserveList  $reference_pclist $iplist
}



if($machine_delete_do -eq 1){
    ##�폜�����J�n##
    Write-Output "���@��̍폜�������n�߂܂�"
    GetPCReserveList $machine_delete_pclist $iplist #�Ώے[���̗\���
    CsvImportCheck $machine_delete_pclist
   
    #�I���v�����v�g�̕\��
    $answer = $host.ui.PromptForChoice("<���s�m�F>","���s���܂����H",$choice,0)

    #$execute = Read-Host "�폜�����s���܂���? y/n"
    Write-Output `n

    if ($answer -eq "0"){
        if($csvinfo_check_flag -eq 0){
            RemoveReserveIP #�\��IP�̍폜
        }
        Write-Output `n
        Write-Output "���|���V�[����MAC�A�h���X���폜���܂�"
        RemovePolicyMacAddress #�e�X�R�[�v�̃|���V�[����Ώے[��MAC�A�h���X���폜����
        Write-Output `n
        Write-Output "�������I�����܂���"
        Write-Output `n
    }else{
        Write-Output "�폜���L�����Z������܂���"
    }
    Write-Output `n
   CheckPolicyMacNo
    Write-Output `n
}




if($add_do -eq 1){
    ##�ǉ������J�n##
    Write-Output "��IP�̒ǉ��������n�߂܂�"
    $global:csvinfo_check_flag = 1
    CsvImportCheck $add_pclist


    #�I���v�����v�g�̕\��
    $answer = $host.ui.PromptForChoice("<���s�m�F>","���s���܂����H",$choice,0)


    if ($answer -eq "0"){
        

            foreach($eh in $exist_host){

                if($eh.ipaddress -ne  $iplist ){
                    try{
                        $eh | Add-DhcpServerv4Reservation -computername $servername -type dhcp -WhatIf
                        $eh | Add-DhcpServerv4Reservation -computername $servername -type dhcp 
                        $global:IPlist_fix_check_flag =1
                    }catch{
                        $temp_text = "�\��IP�A�h���X�ǉ��Ɏ��s���܂����B"
                        Write-Output $temp_text
                        pause
                    }
                }else{
                    $temp_text = $eh.ipaddress + "�͂��łɓo�^������܂�"
                    write-output $temp_text
                    pause
                }
                }

            
        

        Write-Output `n
        Write-Output "���|���V�[��MAC�A�h���X��ǉ����܂�"
        AddPolicyMacAddress #�e�X�R�[�v�̃|���V�[����Ώے[��MAC�A�h���X���폜����
        Write-Output `n
        Write-Output "�������I�����܂���"
        Write-Output `n
    }else{
        Write-Output "�ǉ����L�����Z������܂���"
    }
    Write-Output `n
   CheckPolicyMacNo
    Write-Output `n
}

if($IPdelete_do -eq 1){
    ##�폜�����J�n##
    Write-Output "��IP�̍폜�������n�߂܂�"
    GetPCReserveList_IP $IPdelete_pclist $iplist
   
    $global:all_macaddresslist = $exist_host | get-DhcpServerv4Policy -computername $servername -Name Allow | Select-Object Name,ScopeID,@{n='MacAddress';e={$_.MacAddress}} #�Ώے[����MAC�A�h���X�擾

    #�I���v�����v�g�̕\��
    $answer = $host.ui.PromptForChoice("<���s�m�F>","���s���܂����H",$choice,0)

    #$execute = Read-Host "�폜�����s���܂���? y/n"
    Write-Output `n

    if ($answer -eq "0"){
        if($csvinfo_check_flag -eq 0){
            RemoveReserveIP #�\��IP�̍폜
        }
        Write-Output `n
        Write-Output "���|���V�[����MAC�A�h���X���폜���܂�"
        RemovePolicyMacAddress #�e�X�R�[�v�̃|���V�[����Ώے[��MAC�A�h���X���폜����
        Write-Output `n
        Write-Output "�������I�����܂���"
        Write-Output `n
    }else{
        Write-Output "�폜���L�����Z������܂���"
    }
    Write-Output `n
   CheckPolicyMacNo
    Write-Output `n
}

pause




