# PATHedit

A PowerShell GUI tool for managing Windows PATH environment variables.

## Features

- View and edit System, User, and Temporary PATH variables
- Add, edit, delete, and reorder PATH entries
- Visual indication of existing/non-existing paths
- Automatic backup before changes
- Administrator privileges handling
- Double-click to edit entries
- Browse folders support

### Todo

- [ ] Proper import/export functionality
- [ ] Implementing path validation
- [ ] Ability to move blocks of more than one
- [ ] Improve backup management

## Installation

### Option 1: Direct Installation
```powershell
irm https://raw.githubusercontent.com/fl4pj4ck/pathedit/main/src/pathedit.ps1 | iex
```

### Option 2: Manual Installation
1. Clone the repository or download the latest release
2. Run `src/pathedit.ps1`

## Usage

1. Launch `pathedit.ps1` - it will request administrator privileges if needed
2. Use the tabs to switch between System, User, and Temporary PATH variables
3. Use buttons or double-click to modify entries
4. Click OK to apply changes (backup will be created automatically)

## Requirements

- Windows PowerShell 5.1 or PowerShell Core 7.x
- Windows OS
- Administrator privileges (for System PATH modifications)

## License

MIT License - see LICENSE file for details
