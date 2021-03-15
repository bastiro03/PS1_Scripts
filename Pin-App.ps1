function Pin-App {
    param(
        [parameter(mandatory=$true)][ValidateNotNullOrEmpty()][string[]]$appname,
        [switch]$unpin
    )
    $actionstring = @{$true='Von "Start" lösen|Unpin from Start';$false='An "Start" anheften|Pin to Start'}[$unpin.IsPresent]
    $action = @{$true='unpinned from';$false='pinned to'}[$unpin.IsPresent]
    $apps = (New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | ?{$_.Name -in $appname}
    
    if($apps){
        $notfound = compare $appname $apps.Name -PassThru
        if ($notfound){write-error "These App(s) were not found: $($notfound -join ",")"}

        foreach ($app in $apps){
            $appaction = $app.Verbs() | ?{$_.Name.replace('&','') -match $actionstring}
            if ($appaction){
                $appaction | %{$_.DoIt(); return "App '$($app.Name)' $action Start"}
            }else{
                write-error "App '$($app.Name)' is already pinned/unpinned to/from start or action not supported."
            }
        }
    }else{
        write-error "App(s) not found: $($appname -join ",")"
    }
}