# KnotDo Support Tools

Support scripts for [KnotDo](https://knotdo.app) — the offline-first task manager.

---

## Scripts

### `export-ms-todo.ps1` — Microsoft To-Do Export via Graph API

Exports all your Microsoft To-Do lists and tasks to a JSON file, preserving list names, due dates, notes, priorities, and completion status. Handles 100k+ tasks.

**No Azure app registration required.** Uses Microsoft's own first-party Graph Command Line Tools client with device code authentication.

#### Requirements

- Windows with PowerShell 5.1+ (built into Windows 10/11)
- Internet connection
- A Microsoft account with To-Do tasks

#### Usage

**Option A — Download and run directly:**

1. Download [`export-ms-todo.ps1`](./export-ms-todo.ps1)
2. Open PowerShell and run:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   .\export-ms-todo.ps1
   ```
3. The script installs the `Microsoft.Graph.Tasks` module on first run (this takes about a minute)
4. You will see a prompt like:
   ```
   To sign in, use a web browser to open the page https://microsoft.com/devicelogin
   and enter the code ABCD12345 to authenticate.
   ```
5. Open [microsoft.com/devicelogin](https://microsoft.com/devicelogin), enter the code, and sign in with the Microsoft account that has your To-Do data
6. The export saves to `Downloads\ms-todo-export.json`

**Option B — From the KnotDo project folder:**

```powershell
.\scripts\export-ms-todo.ps1
```

#### Importing into KnotDo

1. Go to your KnotDo app → **Import**
2. Choose **Microsoft Graph Export**
3. Upload `ms-todo-export.json`
4. KnotDo imports all lists and tasks with full list names intact

#### What gets exported

| Field | Exported |
|-------|----------|
| List names | Yes |
| Task title | Yes |
| Status (todo / in progress / done) | Yes |
| Priority (low / normal / high) | Yes |
| Due date | Yes |
| Completion date | Yes |
| Notes / body | Yes |

#### Troubleshooting

**"execution of scripts is disabled"**
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

**Module install fails**
```powershell
Install-Module Microsoft.Graph.Tasks -Scope CurrentUser -Force -AllowClobber
```

**Wrong account signs in**
Close all browser windows, go to [microsoft.com/devicelogin](https://microsoft.com/devicelogin) in a private/incognito window, and enter the code there.

---

## License

MIT
