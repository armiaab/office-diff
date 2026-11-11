# office-diff

See what changed between two versions of a **Word** or **PowerPoint** file, using Microsoft
Office's own tools. Works inside git (`git diff`, `git compare`) and without git (drag & drop).

| File type | `git diff` / `git wdiff` (terminal) | `git compare` / `compare.cmd` (window) |
|---|---|---|
| `.docx` `.docm` | text of every paragraph, table, equation, footnote | **Word Compare**: old vs new as tracked changes |
| `.pptx` `.pptm` | text of every slide + speaker notes, in slide order | both versions **side by side** in PowerPoint (old left, new right) + text diff |
| `.xlsx` `.xlsm` | binary | **Spreadsheet Compare** (Office Pro Plus / Microsoft 365 Apps) |
| `.doc` `.ppt` `.xls` | binary | same as above (no text diff for `.ppt`) |

> PowerPoint's own *Compare & Merge* was retired in Microsoft 365 version 2502 (Microsoft removed
> it and its `Merge`/`MergeWithBaseline` automation calls), so side by side + text diff is the
> closest built-in option left.

## Requirements

- Windows with [Git for Windows](https://git-scm.com/download/win)
- Microsoft Word / PowerPoint desktop apps (Microsoft 365 or Office 2016+)
- Windows PowerShell 5.1 (built in). PowerShell 7 (`pwsh`) is used for `git diff` when installed - it's ~2x faster.

## Install

```powershell
git clone <this repo> C:\Users\<you>\source\repos\office-diff
cd C:\Users\<you>\source\repos\office-diff
powershell -ExecutionPolicy Bypass -File install.ps1                # every repo of your user
powershell -ExecutionPolicy Bypass -File install.ps1 -Repo C:\thesis # or: just one repo
```

The installer writes `office-diff.gitconfig` (machine-specific, git-ignored) and adds it to your
git config with `include.path`, then marks `*.docx`/`*.pptx` as `diff=office` in git's attributes
file. Re-run it if you move this folder.

Uninstall: `install.ps1 -Uninstall` (add `-Repo <path>` if you installed for one repo).

## Use

```powershell
git diff                                  # text changes inside .docx/.pptx since the last commit
git wdiff                                 # same, word by word - best for long paragraphs
git log -p -- thesis.docx                 # text history of one document
git compare                               # Office compare: working files vs last commit
git compare HEAD~1 HEAD -- thesis.docx    # any two commits
git compare --no-index draft1.pptx draft2.pptx   # two separate files, no history needed
```

Without git: drag two files onto `compare.cmd` (the older one by modified date is treated as the
old version). For a right-click menu entry, put a shortcut to `compare.cmd` in `shell:sendto`, then
select two files > right-click > *Send to* > the shortcut.

The scripts also run on their own:

```powershell
.\office-compare.ps1 old.docx new.docx
.\office-textconv.ps1 slides.pptx          # print the text git diff sees
```

## How it works

- `office-textconv.ps1` - git *textconv* driver. Reads the Office Open XML inside the file (no Office
  needed) and prints one line per paragraph; pictures show up as `[image]`. Results for committed
  versions are cached (`cachetextconv`), so `git log -p` stays fast.
- `office-compare.ps1` - git *difftool*. Copies both versions to `%TEMP%\office-diff\` (originals are
  never opened or locked; copies older than a day are cleaned up) and drives Word / PowerPoint through
  COM automation. Non-Office files fall back to a normal `git diff`.

## Troubleshooting

- **`git compare` does nothing for a file** - it must be modified relative to what you compare against; check with `git status`.
- **Word/PowerPoint stays in the background** - click its taskbar icon; Windows sometimes blocks focus stealing.
- **Script blocked by execution policy** - always run via `powershell -ExecutionPolicy Bypass -File ...` (the git config already does this).
