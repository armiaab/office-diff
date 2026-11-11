# office-diff

See what changed between two versions of a **Word** or **PowerPoint** file, using Microsoft
Office's own tools. Works inside git (`git diff`, `git compare`) and without git (drag & drop).

| File type | `git diff` / `git wdiff` (terminal) | `git compare` / `compare.cmd` (window) |
|---|---|---|
| `.docx` `.docm` | the document as text: headings, paragraphs, tables, text boxes, equations, headers and footers, footnotes, endnotes, comments | **Word Compare**: old vs new as tracked changes |
| `.pptx` `.pptm` | text of every slide and its speaker notes, in slide order | **Slide report** in your browser: changed, added, removed and moved slides, old and new side by side, with word-level text changes |
| `.xlsx` `.xlsm` | binary | **Spreadsheet Compare** (Office Pro Plus / Microsoft 365 Apps for enterprise) |
| `.doc` `.ppt` `.xls` | binary | same as above |

PowerPoint's own *Compare & Merge* was retired in Microsoft 365 version 2502, so office-diff builds
its own comparison: PowerPoint renders every slide of both versions, slides are matched by
PowerPoint's internal slide ID (so reordering is recognised), and anything that looks different or
has different text or notes is flagged.

## Requirements

- Windows with [Git for Windows](https://git-scm.com/download/win)
- Microsoft Word / PowerPoint desktop apps (Microsoft 365 or Office 2016+) for `git compare`;
  `git diff` reads the files directly and works without Office
- Windows PowerShell 5.1 (built in). PowerShell 7 (`pwsh`) is used for `git diff` when installed - it starts faster.

## Install

```powershell
git clone <this repo> C:\Users\<you>\source\repos\office-diff
cd C:\Users\<you>\source\repos\office-diff
powershell -ExecutionPolicy Bypass -File install.ps1                # every repo of your user
powershell -ExecutionPolicy Bypass -File install.ps1 -Repo C:\thesis # or: just one repo
```

The installer writes `office-diff.gitconfig` (machine-specific, git-ignored) and adds it to your
git config with `include.path`, then marks `*.docx`/`*.pptx` as `diff=office` in git's attributes
file. **Re-run it after moving or updating this folder**: it repoints git at the new location and
makes git drop text it cached from the previous version of the scripts.

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

### Reading the text diff

`git diff` shows a plain-text rendering of each document:

```text
# Chapter heading                 headings are marked with # (## for level 2, ...)
| Name | Score |                  one line per table row
Picture:[image]                   pictures, charts, SmartArt and media show as [image], [chart], ...
=== Slide 4 (hidden) ===          slide headers say when a slide is hidden
--- notes ---                     speaker notes follow the slide text
=== Comments ===                  Word comments, prefixed with their author
```

Tracked changes are shown as if accepted, so the diff matches what the document says now.

## How it works

- `office-textconv.ps1` - git *textconv* driver. Reads the Office Open XML inside the file (no Office
  needed). Results for committed versions are cached (`cachetextconv`), so `git log -p` stays fast.
- `office-compare.ps1` - git *difftool*. Copies both versions to `%TEMP%\office-diff\` (originals are
  never opened or locked; copies older than a day are cleaned up) and drives Word / PowerPoint through
  COM automation. Word documents that contain tracked changes are compared as if those changes were
  accepted, which is what Word itself does. Non-Office files fall back to a normal `git diff`.
- `slide-report.html` - the page template for the PowerPoint slide report. Press <kbd>J</kbd> / <kbd>K</kbd>
  to step through slides, use the buttons at the top to show or hide each kind of change.

## Troubleshooting

- **`git compare` does nothing for a file** - it must be modified relative to what you compare against; check with `git status`.
- **Every slide shows as changed** - the two decks share no slide IDs (for example, one was rebuilt from scratch), so slides are matched by position instead.
- **Word stays in the background** - click its taskbar icon; Windows sometimes blocks focus stealing.
- **Script blocked by execution policy** - always run via `powershell -ExecutionPolicy Bypass -File ...` (the git config already does this).
