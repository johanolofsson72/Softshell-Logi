# Softshell Logi

A Windows shell replacement written in Visual Basic 6 around the year 2000. Instead of the Explorer desktop and taskbar, Softshell Logi draws its own taskbar with a start menu, a quick-launch menu, a clock, and a list of running windows. The `.vbp` project file dates the build to `2000-06-01` and labels itself "Shellreplacement". This is an old hobby project, archived here as-is.

The name is the project's own — "Logi" is not logistics or lodging, it is just what the author called the shell.

## What it does

On startup (`Sub Main` in `modStartUp.bas`) it hides the Explorer taskbar and shows its own `frmTaskbar` form. From there:

- **Start menu** — reads the real Windows special folders (Start Menu, Desktop, Favorites, My Documents) via `SHGetFolderPath` in `modGetSpecialFolderLocation.bas`, then builds a cascading menu from those folders (`frmMenuSystem`, `frmSubmenu`).
- **Quick menu** — a second pop-up menu driven by a `Quickmenu` folder of shortcuts.
- **Taskbar** — `frmTaskbar` enumerates open windows into a task list, minimising/restoring them on click, and updates a date/time display on a timer.
- **Launching** — opens files and shortcuts through `ShellExecute` (`modShellExec.bas`).
- **Shutdown / restart** — `frmAutoExitWindows` and a `Shutdown` (`ExitWindowsEx`) call.
- **Shell swap** — `frmSwap` lets you switch between running Softshell alongside Explorer or as the system shell, using an INI/registry flag (`modErrFix.bas` reads and writes `system.ini`-style settings).
- **Window effects and sound** — `modFormEffects.bas` slides menus up/down with a timer-based delay; `modSound.bas` plays `.wav` files (hover, select, open) through `sndPlaySound`.

Most of the Windows integration is done by declaring Win32 API functions directly (`user32`, `shell32`, `gdi32`, `kernel32`). `modDevice1.bas` holds the bulk of those declarations.

## Tech stack

- **Visual Basic 5.0 / 6.0** — the project targets VB6 (`Type=Exe`, output `Logi.exe`).
- **Win32 API** — direct `Declare` calls for window management, icons, shell folders, and INI files.
- **OCX/ActiveX controls:** `AMCLABEL.OCX` (a label control, not written by the author — see the credit below), `MSCOMCTL.OCX`, and `Comdlg32.ocx`.

## Project structure

```
soft_shell/
  Softshell Logi/              VB6 project
    Softshell Logi Beta 1.0.vbp   project file
    frmTaskbar.frm                taskbar, clock, running-window list
    frmMenuSystem.frm             start menu
    frmSubmenu.frm                submenus
    frmControl.frm / frmSLControl.frm   settings/control UI
    frmSwap.frm                   swap shell vs. run with Explorer
    frmAutoExitWindows.frm        shutdown/restart
    mod*.bas                      Win32 declarations, startup, sound, effects, INI, shell exec
    Icons/  Sound/  Quickmenu/    resources
  This First/
    AMCLabel.ocx                  third-party control
    Read me.txt                   install note
ryska softshell logi.pdf          Russian write-up (Planet Source Code listing)
*.bmp / *.gif                     screenshots and the contest-winner image
vb_soft_shell.htm                 archived listing page
```

`SoftshellLogi1.2ContestWinner.bmp` and `vb_soft_shell.htm` indicate this was submitted to Planet Source Code (the `.htm` is the Russian-language listing).

## Getting started

This needs a Windows machine with the Visual Basic 6 IDE (or VB5). It will not build or run on modern toolchains without that environment, and as a shell replacement it expects a classic Win9x/2000-era Windows.

1. Per `soft_shell/This First/Read me.txt`, copy `AMCLabel.ocx` into the Windows `\system` directory and register it.
2. Open `soft_shell/Softshell Logi/Softshell Logi Beta 1.0.vbp` in VB6.
3. Make sure the referenced controls (`MSCOMCTL.OCX`, `Comdlg32.ocx`, `AMCLABEL.OCX`) are present and registered.
4. Run from the IDE, or compile to `Logi.exe`.

To use it as the actual system shell rather than alongside Explorer, the `frmSwap` screen writes the relevant shell setting and restarts Windows. Treat that as a one-way trip on a throwaway/VM install.

## Credits

The `AMCLabel.ocx` control was not written by the project author (noted in `Read me.txt` and the comment headers). Other comments credit "Brian" for the sound and several API modules, "Jeffrey C Tatum" for the INI read/write routines, and Microsoft sample code for parts of `modDevice1.bas`. Original author: Johan Olofsson (softworld / softworlddata.com).

## Status

Archived. Last meaningful work on the code is from around 2000; the repo was last touched in October 2024 only to collect the files. It is kept here for reference and nostalgia, not maintained. Expect it to run only inside a period-appropriate Windows environment with VB6 installed.
