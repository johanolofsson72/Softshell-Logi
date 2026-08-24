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
ryska softshell logi.pdf          Russian write-up (Ishodnik.Ru listing, PDF)
SoftshellLogi1.2ContestWinner.bmp Planet Source Code listing for Beta 1.2, with the CONTEST WINNER badge
Softshell at Shell City.bmp       shellcity.net "Daily Update", 25 September 2000
softshelllogi.bmp                 Beta 1.0 running as the system shell, "(c) Softworld (tm)" in every menu
vb_soft_shell.htm                 archived Ishodnik.Ru listing (Russian)
```

`SoftshellLogi1.2ContestWinner.bmp` and `vb_soft_shell.htm` document where this was published and how it was received — see **Publication and reception** below.

## Publication and reception

Softshell Logi was released publicly in 2000, with full source, and was picked up in at least three places. The screenshots and the archived page in this repo are the surviving evidence.

### Planet Source Code — two separate entries

| | Beta 1.2 | Beta 1.3 |
|---|---|---|
| Submitted | **2000-08-01, 9:23:47 PM** | **2000-09-03, 6:39:44 PM** |
| Code ID | *(not recorded)* | **11233** |
| Level | Advanced | Advanced |
| User rating | rated by **19 users** | 4 of 5, rated by **15 users** |
| Accesses / downloads | **1,236** *(as shown in the contemporaneous screenshot)* | **18,830** *(as listed later)* |
| Award | **CONTEST WINNER** | — |

`SoftshellLogi1.2ContestWinner.bmp` is a screenshot of the **Beta 1.2** listing and carries the laurel "CONTEST WINNER" badge. Note that the award sits on the 1.2 entry, not on 1.3 — anyone checking code ID 11233 is looking at 1.3 and will not find it. The same listing displays a Brainbench "Certified Professional — Visual Basic 6.0 Programmer" badge for the author.

The two access counts are not in conflict: they are different entries, counted at different times, roughly a quarter-century apart.

### Shell Extension City (shellcity.net)

`Softshell at Shell City.bmp` is a screenshot of that site's **DAILY UPDATE for September 25, 2000**, where Softshell is the lead item, above `MMDESK` and `ONE BUTTON`. The write-up is the author's own description, reproduced in the first person, and the site pointed readers to Planet Source Code for the free download rather than hosting it.

### Исходник.Ру / Ishodnik.Ru (Russia)

`vb_soft_shell.htm` is **not** a Planet Source Code page. It is a listing from **Исходник.Ру — "Сайт профессионального программиста"** ("Source.Ru — the professional programmer's site"), a Russian source-code library of the era covering C/C++, Pascal, Delphi, Kylix, Visual Basic, Assembler, WAP and PalmOS. `ryska softshell logi.pdf` is the same listing as a PDF.

The page credits **"Автор: Johan Olofsson"** and links to the author's own site, `softworlddata.com`. It describes the program as a movable Windows shell that reproduces the taskbar and adds cascading menus with Linux-style on-the-fly menu additions and scrolling, and singles out for praise that it ships with complete source: *"А самый главный плюс этой проги, то что она с полным исходным кодом"* — "the main advantage of this program is that it comes with full source code."

### Branding

The running program renders **© Softworld ™** in the footer of every menu panel (visible in `softshelllogi.bmp`, which shows Beta 1.0 running as the actual system shell). The `.vbp` carries `VersionCompanyName = "softworld"`, and Ishodnik.Ru links the author to `softworlddata.com`.

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
