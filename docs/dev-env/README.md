# VS Code Development Environment

These files capture the VS Code extensions and user settings from the original VM.

## Install Extensions
From the cloned `watermark-python` repo on the laptop:

```powershell
Get-Content .\docs\dev-env\vscode-extensions.txt | ForEach-Object { code --install-extension $_ }
```

## Restore Settings
This overwrites the laptop's current VS Code user settings:

```powershell
Copy-Item .\docs\dev-env\vscode-settings.json "$env:APPDATA\Code\User\settings.json" -Force
```

## Notes
- No `keybindings.json` existed on the original VM, so there are no custom keybindings to restore.
- This does not transfer account sign-ins. Re-authenticate GitHub, ChatGPT/Codex, Azure, and related extensions on the laptop.
