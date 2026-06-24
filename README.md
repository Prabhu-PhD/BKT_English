# BKT Toolbox — English Edition

This is an **English-translated build** of the **Business Kasper Toolbox (BKT)**, a powerful free add-in that extends Microsoft PowerPoint (and optionally Excel and Visio) with dozens of productivity features — shape alignment, formatting tools, agendas, shape tables, Harvey balls, and much more.

The original toolbox is developed **in German**. This fork translates the user-facing interface (ribbon labels, tooltips, and dialogs) into English so English-speaking teams can use it comfortably.

> **Quick install:** download **[`_releases/bkt_install_english_v3.1.0.exe`](_releases/)**, run it, restart PowerPoint. See **[HOW-TO-INSTALL.txt](HOW-TO-INSTALL.txt)** for step-by-step instructions you can forward to colleagues.

---

## 🙏 Credits — all original work belongs to the pyro-team

The Business Kasper Toolbox and Framework were created and are maintained by the **[pyro-team](https://github.com/pyro-team)**:

- **Original project:** https://github.com/pyro-team/bkt-toolbox
- **Website:** https://www.bkt-toolbox.de
- **Documentation / Wiki:** https://github.com/pyro-team/bkt-toolbox/wiki
- **License:** MIT (see [LICENSE.md](LICENSE.md)) — retained unchanged.

All credit for the toolbox itself — its architecture, features, and years of work — goes to the original authors. This fork contributes **only an English translation of the interface** and packaging; it adds no new functionality. If you find BKT useful, please support and star the **[upstream project](https://github.com/pyro-team/bkt-toolbox)**.

---

## What was translated

The translation covers the text users actually see:

- **Ribbon UI** in the Python sources — the `label`, `supertip`, `screentip`, `title`, and `description` attributes of ribbon controls.
- **WPF dialogs** in the `.xaml` files — the `Content`, `Header`, `Text`, `ToolTip`, and `Title` attributes.

In total, **2,706 unique German strings** were translated (≈2,750 occurrences across 88 files). Code comments and a few developer-only/test strings were intentionally left as-is.

**Nothing functional was changed.** The translation never touches control IDs (`id=`), Office image references (`image_mso=`/`idMso=`), callbacks (`on_action=`, …), or WPF data bindings (`{Binding …}`) — only the human-readable display text. The prebuilt `bin/BKT.dll` is **unchanged**: BKT runs on IronPython and loads the `.py`/`.xaml` files at runtime, so **no C# recompilation is required**.

### The translation pipeline (`_translation/`)

The translation is reproducible, not a one-off hand-edit:

| Script | Purpose |
|---|---|
| `extract.py` | Scans `features/`, `bkt/`, `modules/` and collects every translatable string into `strings.json`. |
| `make_worklist.py` | Produces `keys.json` (source of truth) + an indexed `worklist.tsv`. |
| `parts/*.tsv` | The English translations, keyed by index so the originals are never re-typed (avoids whitespace/Unicode mismatches). |
| `merge.py` | Maps the English back to the exact original keys → `translations.json`. |
| `apply.py` | Performs a targeted, attribute-scoped in-place replacement. Supports `--root <dir>` to translate any install. |

**To re-apply after pulling an upstream update:** re-run `extract.py`, translate any new strings into a new `parts/NNN.tsv`, then `merge.py` and `apply.py`.

---

## Problems faced & how they were resolved

This translation surfaced a few non-obvious issues worth recording:

1. **No internationalization layer.** BKT has no gettext/`.po` system — German is hard-coded throughout. *Resolution:* an extract→translate→apply pipeline scoped strictly to display attributes, so functional identifiers can never be corrupted.

2. **Encoding & escaping.** Sources are UTF-8 and represent in-string line breaks as the literal two-character escape `\n` inside single-line string literals. A naïve replace would corrupt these (or break a single-quoted literal that contains an apostrophe, e.g. *Python's*). *Resolution:* `apply.py` preserves the literal `\n`, keeps each attribute's original quote style, and escapes the delimiter quote when needed. Correctness was checked with a **differential compile** (every changed `.py` compiled before vs. after — zero newly-broken files) and an XML well-formedness check on all 40 `.xaml` files.

3. **"The add-in won't load" — a testing red herring.** During verification, PowerPoint was being launched through an automation tool that does **not** re-process newly-registered Office add-ins (they never reach `…\PowerPoint\AddInLoadTimes`). This produced a long chain of false "doesn't load" conclusions and several wrong theories (org add-in policy; modern Office blocking legacy COM add-ins). *Resolution:* launching PowerPoint **normally** processes the add-in immediately — the translated engine loads and renders the English ribbon. **Lesson: always verify add-in loading with a normal launch**, and confirm via `AddInLoadTimes` + the `bkt-debug-py.log`.

4. **A VSTO wrapper experiment (abandoned).** Before the above was understood, a thin **VSTO** host shim was built (`dotnet/bkt-vsto/`) to load the BKT engine via the VSTO runtime. It compiles, signs, registers, and is recognized by PowerPoint — but it **crashes during BKT's IronPython bootstrap** (a VSTO-hosting/threading incompatibility). It is kept for reference but is **not used**; the standard COM add-in is the working path. *(Its private signing certificate is intentionally excluded from this repo.)*

5. **Packaging for a team.** *Resolution:* the project's own Inno Setup script was duplicated as **`setup/innosetup_english.iss`** with English wizard text, producing a per-user, no-admin installer (`_releases/bkt_install_english_v3.1.0.exe`) that bundles the translated tree and registers the add-in automatically.

---

## Installation

**For most users (recommended):** run **`_releases/bkt_install_english_v3.1.0.exe`** and restart PowerPoint. Full step-by-step: **[HOW-TO-INSTALL.txt](HOW-TO-INSTALL.txt)**.

**For developers (from source):** clone the repo and run `installer\install.bat`. After an upstream update, the file may need to be re-run. The add-in is per-user (HKCU) and registers automatically; remove it with `installer\uninstall.bat`.

To build the English installer yourself (requires [Inno Setup 6](https://jrsoftware.org/isinfo.php)):

```cmd
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup\innosetup_english.iss
```

Output lands in `_releases\`.

---

## System requirements

Windows with Office 2013+ (the simple installer targets 2013 and newer). A Mac version is not possible because Microsoft does not expose the COM add-in interface in Mac Office. Requires .NET Framework 4.5+.

> The toolbox is active in **PowerPoint** by default; it is also available in Excel, Word, Outlook, and Visio via *File → Options → Add-Ins*.

---

## Built on (from the original project)

- [IronPython](https://github.com/IronLanguages/ironpython2) · [Fluent.Ribbon](https://github.com/fluentribbon/Fluent.Ribbon) · [ControlzEx](https://github.com/ControlzEx/ControlzEx) · [MahApps.Metro](https://github.com/MahApps/MahApps.Metro) · [MouseKeyHooks](https://github.com/gmamaladze/globalmousekeyhook) · [InnoSetup](http://www.jrsoftware.org/isinfo.php) · [Google Material Icons](https://material.io/tools/icons/) & [Material Design Icons](https://materialdesignicons.com/)

---

*This is an unofficial, community translation. It is not affiliated with or endorsed by the pyro-team. For the canonical, maintained, German-language toolbox, please use the [original repository](https://github.com/pyro-team/bkt-toolbox).*
