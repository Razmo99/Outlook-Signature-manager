param (
    [string]$signature_name,
    [switch]$force = $false
)

#region Config
$domain_fqdn = "contoso.local"
$script_version = "0.0.0.002"
$running_as_user = [System.Environment]::userName
$form_icon = ".\update_outlook_icon.ico"

$form_name = "Contoso"
# Path to signature templates & versions
$network_path = "\\$domain_fqdn\NETLOGON\Signature\Templates"
$user_signatures_path = "$env:APPDATA\Microsoft\Signatures"
$log_path = "$env:USERPROFILE\update_outlooksignature_logs"
#endregion Config

#region Start Logging
function invoke-rotating_log_file_handler {
    [OutputType([void])]
    param (
        [Parameter(Mandatory = $true)][string]$path,
        [Parameter(Mandatory = $true)][string]$name,
        [Parameter(Mandatory = $false)][int]$backup_count = 10
    )

    $time_stamp = Get-Date -Format yyyy-MM-dd
    $log_file_name = '{0}_{1}.log' -f $name, $time_stamp
    $log_file = Join-Path -Path $path -ChildPath $log_file_name
    #Change this value to how many log files you want to keep
    If (Test-Path -Path $path) {
        #Make some cleanup and keep only the most recent ones
        $filter = '*_????-??-??.log'
        Get-ChildItem -Path $filter |
        Sort-Object -Property LastWriteTime -Descending |
        Select-Object -Skip $backup_count |
        Remove-Item -Verbose
    }
    else {
        #No logs to clean but create the Logs folder
        New-Item -Path $path -ItemType Directory -Verbose
    }
    return $log_file
}
function write-log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$path,
        [Parameter(Mandatory = $false)]
        [ValidateSet("FATAL", "ERROR", "WARN", "INFO", "DEBUG", "TRACE")]
        [string]$log_level = "INFO",
        [Parameter(Mandatory = $true)]
        [string]$msg
    )
    begin {
        # Checks our our path ends with .log or .txt case insensitive
        if ($PSBoundParameters.ContainsKey('Path')) {
            if (!([regex]::match($PSBoundParameters['Path'].toLower(), "(.*\.(log|txt))")).Success) {
                Write-Warning -Message "Log file path does not end with .log or .txt"
            }

        }
    }
    process {
        $time_stamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssK"
        $path_exists = Test-Path $path -ErrorAction Stop
        if (!$path_exists) {
            [void](New-Item -Path $path -ItemType File -ErrorAction Stop -WhatIf:$false)
        }
        $msg = "$($time_stamp) - $($log_level) - $($msg)"
        Add-Content -Path $path -Value $msg -WhatIf:$false
    }
    end {}
}

$log_file = invoke-rotating_log_file_handler -path $log_path -name $running_as_user
write-log -path $log_file -log_level DEBUG -msg "script version: $script_version"
#endregion Start Logging

#region Pre Checks
# Create a query to lookup the user the script is running as
[ADSISearcher]$ad_lookup = "samaccountname=$running_as_user"
$user = $ad_lookup.FindOne()

# Exit if we are unable to find the user
If ($null -eq $user) {
    write-log -path $log_file -log_level ERROR -msg "Failed to Lookup user: $running_as_user"
    Exit 1
}
#endregion Pre Checks
#region Check user has the newest signature template

# Get the latest template from the network  (we are expecting these to start with VXX)
# Hence we sort by name and pick the last which should be the newest
# I.E
# V1_contoso.local
# V2_contoso.local
$network_template = Get-ChildItem -Directory $network_path |
Where-Object { $_.Name -like "*$($signature_name)*" } |
Sort-Object Name |
Select-Object -Last 1
# Signautures on the local workstation will be appended with _files
if ($network_template) {
    $network_template_name = $network_template.name -replace '_files'
}
else {
    write-log -path $log_file -log_level WARN `
        -msg "Failed to retreive network template or match a template to the running users UPN suffix"
    exit 1
}


# Get the latest template from the local workstaion we are expecting these to start with VXX
$workstation_template = Get-ChildItem -Directory $user_signatures_path |
Where-Object { $_.Name -like "*$($signature_name)*" } |
Sort-Object Name |
Select-Object -Last 1

$workstation_template_name = $null
if ($workstation_template) {
    $workstation_template_name = $workstation_template.name -replace '_files'
}


# If script has been run with the -force parameter ignores if these is an existing version with the same name
if ($workstation_template_name -eq $network_template_name -and $force -eq $true) {
    write-log -path $log_file -log_level WARN -msg "Workstation and Network Signature have the same name, force parameter ignored exiting"
    exit 0
}
# Workstation version is already up to date no action needed
if ($workstation_template_name -eq $network_template_name) {
    write-log -path $log_file -log_level INFO -msg "Workstation and Network Signature have the same name, no action needed"
    exit 0
}
#endregion Check user has the newest signature template

#region form

# Load in required assemblies for forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Flexible Working helper
Function Flex ($parent) {
    [CmdletBinding()]

    $week_days = @(
        "Monday"
        "Tuesday"
        "Wednesday"
        "Thursday"
        "Friday"
    )

    $f_flex = New-Object System.Windows.Forms.Form -Property @{
        Text          = "$form_name Signature Creator - Flexible Working"
        AutoScale     = $true
        AutoSize      = $true
        AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
        StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        font          = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Regular)
        Topmost       = $True
        backcolor     = [System.Drawing.Color]::White
        icon          = $form_icon
    }

    $tlp_flex = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
        parent   = $f_flex
        AutoSize = $true
        dock     = [System.Windows.Forms.DockStyle]::Fill
        Padding  = 15
    }

    $l_flex_description = New-Object System.Windows.Forms.Label -Property @{
        Text      = "Include your flexible working arrangements using the below"
        autosize  = $true
        TextAlign = [System.Drawing.ContentAlignment]::TopCenter
        dock      = [System.Windows.Forms.DockStyle]::Left
        padding   = 5
    }
    $tlp_flex.Controls.Add($l_flex_description, 0, 0)
    $tlp_flex.SetColumnSpan($l_flex_description, 2)

    #####################################################################
    # PART TIME
    #####################################################################
    $l_part_time = New-Object System.Windows.Forms.Label -Property @{
        Text      = "If you work part-time, select`n the days that apply."
        autosize  = $true
        TextAlign = [System.Drawing.ContentAlignment]::TopCenter
        Anchor    = [System.Windows.Forms.AnchorStyles]::None

        padding   = 5
    }
    $tlp_flex.Controls.Add($l_part_time, 0, 1)

    $lb_part_time = New-Object System.Windows.Forms.ListBox  -Property @{
        SelectionMode = [System.Windows.Forms.SelectionMode]::MultiSimple
        autosize      = $true
        padding       = 5
        Anchor        = [System.Windows.Forms.AnchorStyles]::None
    }
    $week_days | ForEach-Object {
        [void]$lb_part_time.Items.Add($_)
    }
    $tlp_flex.Controls.Add($lb_part_time, 0, 2)

    #####################################################################
    # REMOTLEY
    #####################################################################

    $l_remotley = New-Object System.Windows.Forms.Label -Property @{
        Text      = "If you work part-time, select`nthe days that apply."
        autosize  = $true
        TextAlign = [System.Drawing.ContentAlignment]::TopCenter
        Anchor    = [System.Windows.Forms.AnchorStyles]::None
        padding   = 5
    }
    $tlp_flex.Controls.Add($l_remotley, 1, 1)

    $lb_remotley = New-Object System.Windows.Forms.ListBox  -Property @{
        SelectionMode = [System.Windows.Forms.SelectionMode]::MultiSimple
        autosize      = $true
        Anchor        = [System.Windows.Forms.AnchorStyles]::None
        padding       = 5
    }
    $week_days | ForEach-Object {
        [void]$lb_remotley.Items.Add($_)
    }
    $tlp_flex.Controls.Add($lb_remotley, 1, 2)

    #####################################################################
    # OK BUTTON
    #####################################################################

    [scriptblock]$finish_form = {
        If ($lb_part_time.SelectedItems -and $lb_remotley.SelectedItems) {
            $Flex = "Please note I work part-time and my working days are: $($lb_part_time.SelectedItems -join ", ") and I work remotely on: $($lb_remotley.SelectedItems -join ", ")."
        }
        ElseIf ($lb_remotley.SelectedItems) {
            $Flex = "Please note I work remotely on: $($lb_remotley.SelectedItems -join ", ")."
        }
        Elseif ($lb_part_time.SelectedItems) {
            $Flex = "Please note I work part-time and my working days are: $($lb_part_time.SelectedItems -join ", ")."
        }
        $parent.Enabled = $true
        $parent.Text = $Flex
    }

    $btn_flex = New-Object System.Windows.Forms.Button -Property @{
        autosize     = $true
        Text         = "OK"
        dock         = [System.Windows.Forms.DockStyle]::Right
        DialogResult = [System.Windows.Forms.DialogResult]::OK
        Add_Click    = $finish_form
    }
    $tlp_flex.Controls.Add($btn_flex, 1, 3)
    $tlp_flex.SetColumnSpan($btn_flex, 2)

    $f_flex.AcceptButton = $btn_flex
    $f_flex.ShowDialog()

}
#endregion Flexible Working helper

#region Init Form and app layout
[System.Drawing.Font]$d_font = [System.Drawing.SystemFonts]::DefaultFont
function handle-mousewheel{
    $m = [System.Windows.Forms.MouseEventArgs]$_
    $this.SuspendLayout()
    if($m.Delta -gt 0){
        $this.font = [System.Drawing.Font]::new($this.font.name, ($this.font.size +1), [System.Drawing.FontStyle]::Regular)
    }else{
        if($this.font.size -gt $d_font.size){
            $this.font = [System.Drawing.Font]::new($this.font.name, ($this.font.size -1), [System.Drawing.FontStyle]::Regular)
        }
        
    }
    $this.ResumeLayout()
}

$form = New-Object System.Windows.Forms.Form -Property @{
    Text          = "$form_name Signature Creator"
    autosize      = $true
    AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    font          = $d_font
    AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    Topmost       = $True
    backcolor     = [System.Drawing.Color]::White
}
$form.add_MouseWheel({handle-mousewheel})

if (Test-Path $form_icon) {
    $form.icon = [System.Drawing.Icon]::new((resolve-path $form_icon))
}
$form_layout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    parent       = $form
    AutoSize     = $true
    AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    dock         = [System.Windows.Forms.DockStyle]::Fill
    Padding      = 15
}
#endregion Init Form and app layout

$l_description = New-Object System.Windows.Forms.Label -Property @{
    Text     = "To create a new email signature please review and confirm your details below."
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
}
$form_layout.Controls.Add($l_description, 0, 0)
$form_layout.SetColumnSpan($l_description, 2)

#####################################################################
# SIGNATURE
#####################################################################

$l_signature = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Signature:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
}
$form_layout.Controls.Add($l_signature, 0, 1)

$tb_signature = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled  = $false
    Text     = $signature_name
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::Fill
    padding  = 5
}
$form_layout.Controls.Add($tb_signature, 1, 1)

#####################################################################
# FULL NAME
#####################################################################
$l_fullname = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Full Name:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
}
$form_layout.Controls.Add($l_fullname, 0, 2)

$tb_fullname = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled  = $true
    AutoSize = $true
    Text     = $user.Properties.name
    dock     = [System.Windows.Forms.DockStyle]::Fill
    padding  = 5
}
$form_layout.Controls.Add($tb_fullname, 1, 2)

#####################################################################
# Title
#####################################################################
$l_title = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Title:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
    
}
$form_layout.Controls.Add($l_title, 0, 3)

$tb_title = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled  = $true
    AutoSize = $true
    Text     = $user.Properties.title
    dock     = [System.Windows.Forms.DockStyle]::Fill
    padding  = 5
}
$form_layout.Controls.Add($tb_title, 1, 3)
#####################################################################
# Address
#####################################################################
$l_address = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Address:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
}
$form_layout.Controls.Add($l_address, 0, 4)

$tb_address = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled  = $true
    autosize = $true
    Text     = $null
    dock     = [System.Windows.Forms.DockStyle]::Fill
    padding  = 5
}
$form_layout.Controls.Add($tb_address, 1, 4)
$tb_address.Text = "$($user.Properties.streetaddress), $($User.Properties.l -replace '\p{N}') $($user.Properties.st)$($user.Properties.l -replace '[a-z]')"


#####################################################################
# Phone
#####################################################################
$l_phone = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Phone:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
}
$form_layout.Controls.Add($l_phone, 0, 6)

$tb_phone = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled  = $true
    AutoSize = $true
    Text     = $user.Properties.telephonenumber -Replace '(\+61|61)(2|3|8)(\d{4})(\d{4})', '$1$2 $3 $4'
    dock     = [System.Windows.Forms.DockStyle]::Fill
    padding  = 5
}
$form_layout.Controls.Add($tb_phone, 1, 6)
#####################################################################
# Mobile
#####################################################################
$l_mobile = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Mobile*:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
}
$form_layout.Controls.Add($l_mobile, 0, 7)

$tb_mobile = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled  = $true
    AutoSize = $true
    dock     = [System.Windows.Forms.DockStyle]::Fill
    padding  = 5
}
$form_layout.Controls.Add($tb_mobile, 1, 7)

$l_mobile_hint = New-Object System.Windows.Forms.Label -Property @{
    Text     = "*To omit your mobile simply leave this field blank"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::left
    padding  = 5
    Font     = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Italic)
}
$form_layout.Controls.Add($l_mobile_hint, 1, 8)

#####################################################################
# SET AS DEFAULT SIGNATURE
#####################################################################
$cb_default_signature = New-Object System.Windows.Forms.CheckBox -Property @{
    Text     = "Set this as the default signature for all new emails?"
    Checked  = $true
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::Left
}
$form_layout.Controls.Add($cb_default_signature, 0, 9)
$form_layout.SetColumnSpan($cb_default_signature, 2)

#####################################################################
# SET AS DEFAULT REPLY / FORWARD SIGNATURE
#####################################################################
$cb_default_reply_forward = New-Object System.Windows.Forms.CheckBox -Property @{
    Text     = "Set this as the default signature for Reply & forward emails?"
    Checked  = $true
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::Left
}
$form_layout.Controls.Add($cb_default_reply_forward, 0, 10)
$form_layout.SetColumnSpan($cb_default_reply_forward, 2)

#####################################################################
# FLEXIBLE WORKING INFORMATION BUTTON
#####################################################################
$btn_flexible_working = New-Object System.Windows.Forms.Button -Property @{
    autosize     = $true
    AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    dock         = [System.Windows.Forms.DockStyle]::Fill
    ForeColor    = [System.Drawing.Color]::Red
    Text         = "Click here to include flexible working information in your signature"
    Add_Click    = { Flex $tb_flexible_working }
}
$form_layout.Controls.Add($btn_flexible_working, 0, 11)
$form_layout.SetColumnSpan($btn_flexible_working, 2)

#####################################################################
# FLEXIBLE WORKING INFORMATION
#####################################################################
$l_flexible_working = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Flexible Working:"
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::Left
    padding  = 5
}
$form_layout.Controls.Add($l_flexible_working, 0, 12)

$tb_flexible_working = New-Object System.Windows.Forms.TextBox -Property @{
    Enabled     = $true
    Multiline   = $true
    AutoSize    = $true
    dock        = [System.Windows.Forms.DockStyle]::Fill
    padding     = 5
    MinimumSize = [system.drawing.point]::new(1, 100)
}
$form_layout.Controls.Add($tb_flexible_working, 1, 12)

#####################################################################
# Table Layout Panel for OK / CANCEL button
#####################################################################
$tlp_confirm_btn = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    autosize = $true
    dock     = [System.Windows.Forms.DockStyle]::Fill
    Padding  = 15
}
# Adjust colum weighting to be 50 so buttons are in in the middle
[void]$tlp_confirm_btn.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle -Property @{SizeType = [System.Windows.Forms.SizeType]::Percent; Width = 50 }))
[void]$tlp_confirm_btn.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle -Property @{SizeType = [System.Windows.Forms.SizeType]::Percent; Width = 50 }))
$form_layout.Controls.Add($tlp_confirm_btn, 0, 13)
$form_layout.SetColumnSpan($tlp_confirm_btn, 2)

#####################################################################
# OK BUTTON
#####################################################################

$btn_ok = New-Object System.Windows.Forms.Button -Property @{
    Text         = "Okay"
    AutoSize     = $true
    AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    dock         = [System.Windows.Forms.DockStyle]::Right
    DialogResult = [System.Windows.Forms.DialogResult]::OK
}
[void]$tlp_confirm_btn.Controls.Add($btn_ok, 0, 0)

#####################################################################
# CANCEL BUTTON
#####################################################################

$btn_cancel = New-Object System.Windows.Forms.Button -Property @{
    Text         = "Cancel"
    AutoSize     = $true
    AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    dock         = [System.Windows.Forms.DockStyle]::Left
    DialogResult = [System.Windows.Forms.DialogResult]::Cancel
}
[void]$tlp_confirm_btn.Controls.Add($btn_cancel, 1, 0)

$result = $form.ShowDialog()
#endregion form


# User exited the form
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
    write-log -path $log_file -log_level INFO -msg "User exited form"
    Exit 0
}
# Create the signatures dir if it doesnt exist
If (!(Test-Path $user_signatures_path)) {
    write-log -path $log_file -log_level DEBUG -msg "Created signatures folder"
    New-Item -Path $user_signatures_path -ItemType Directory
}
# Copy over the new signature files filter to specificly template files
$new_signature_files = Copy-Item "$($network_path)\$($network_template_name)\*" $user_signatures_path -PassThru -Recurse -Force |
Where-Object { $_.Name -like "*$($signature_name).*" }

# Check if the user entered in a Mobile number
$mobile = $null
if ($tb_mobile.Text) {
    write-log -path $log_file -log_level DEBUG -msg "Mobile specified"
    $mobile = "| Mob: $($tb_mobile.Text)"
}
#region replace op on email signature
# Update the contents of the HTML, RTL & TXT Signature content files with the forms collected information
foreach ($file in $new_signature_files) {
    $file_content = Get-Content $file.fullname -Raw
    $file_content |
    ForEach-Object {
        $_ -replace "@FullName", "$($tb_fullname.Text)" `
            -replace "@Title", "$($tb_title.Text)" `
            -replace "@Address", "$($tb_address.Text)" `
            -replace "@TNumber", "Tel: $($tb_phone.Text)" `
            -replace "@MNumber", "$($mobile)" `
            -replace "@Flex", "$($tb_flexible_working.Text)"
    } |
    Set-Content $file.Fullname
    write-log -path $log_file -log_level INFO -msg "Successfully set signature templates"
}
#endregion replace op on email signature

#region set signature as default
if ($cb_default_signature.Checked -or $cb_default_reply_forward.Checked) {
    
    # Get the users default outlook profile to apply the change too
    $default_profile = (Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook").DefaultProfile
    if ($default_profile) {
        $outlook_profile_path = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\" + $default_profile.Trim() + "\9375CFF0413111d3B88A00104B2A6676\"
        # find the sub folder that has a matching account name to the users email address I.E 
        # HKCU:\\...\9375CFF0413111d3B88A00104B2A6676
        #   \00000001
        #       \Account Name
        #   \00000002
        #       \Account Name
        $sub_key = Get-ChildItem $outlook_profile_path | Where-Object {
        ($_ | Get-ItemProperty).'Account Name' -eq $user.Properties.mail
        }
        # Update / Create the registry Keys
        if ($cb_default_reply_forward.Checked) {
            write-log -path $log_file -log_level INFO -msg "Set default Reply-Forward Signature: $network_template_name"
            $null = New-ItemProperty -Path $sub_key.PSPath -Name 'Reply-Forward Signature' -Value $network_template_name -PropertyType 'String' -Force
        }
    
        if ($cb_default_signature.Checked) {
            write-log -path $log_file -log_level INFO -msg "Set default New Signature: $network_template_name"
            $null = New-ItemProperty -Path $sub_key.PSPath -Name 'New Signature' -Value $network_template_name -PropertyType 'String' -Force
        }
    }
    else {
        write-log -path $log_file -log_level WARN -msg "No outlook default profile found"
    }
}
#endregion set signature as default
