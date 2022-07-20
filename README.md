Open a PowerShell terminal in this directory and run the following commands

```
npx office-addin-dev-certs install
npx http-server . -p 3000 -S --cert $Env:USERPROFILE\.office-addin-dev-certs\localhost.crt --key $Env:USERPROFILE\.office-addin-dev-certs\localhost.key
```
