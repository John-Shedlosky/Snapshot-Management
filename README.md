1. Generate the key (PowerShell)

New-Item -ItemType Directory -Force "$env:USERPROFILE.ssh" | Out-Null

ssh-keygen -t rsa -b 4096 -f "$env:USERPROFILE.ssh\fb_rsa" -C "label-for-the-key"

Get-Content "$env:USERPROFILE.ssh\fb_rsa.pub" | Set-Clipboard


2. Log in with your password and upload the key on FlashBlade

ssh pureuser@flashblade-array-name

pureadmin setattr --publickey pureuser

Enter key: <paste, press Enter>

pureadmin list --publickey pureuser

Because you're logged in as pureuser, you can also leave off the username. The command then sets the key for the current user. 

3. Lock down the private key file in Powershell

icacls "$env:USERPROFILE.ssh\fb_rsa" /inheritance:r /grant:r "$($env:USERNAME):(R)"


4. Test and run commands
ssh -i "$env:USERPROFILE.ssh\fb_rsa" -o BatchMode=yes pureuser@<fb> "purearray list"

ssh -i "$env:USERPROFILE.ssh\fb_rsa" pureuser@<fb> "purefs list --csv" | ConvertFrom-Csv | Format-Table

5. Optional: SSH config shortcut in $env:USERPROFILE.ssh\config

Host fb01
    
    HostName <fb-mgmt-vip-or-fqdn>
    
    User pureuser
    
    IdentityFile ~/.ssh/fb_rsa

After that, run ssh fb01 "purefs list"
