# Developer guidelines

See the [OpenXML specification](https://www.ecma-international.org/publications/standards/Ecma-376.htm) for reference.
In order for XLibur, the wrapper around OpenXML, to support all the features, we rely on community contributions.

Here are some tips.

* Before starting a large pull request, log an issue and outline the problem and a broad outline of your solution. The maintainers will discuss the issue with you and possibly propose some alternative approaches to align with the XLibur development conventions.
* Please submit pull requests that are based on the `main` branch.
* Where possible, pull requests should include unit tests that cover as many uses cases as possible.

## Pull request titles

**The PR title becomes the release note**, so write it as a short imperative summary of the
change. Titles follow [Conventional Commits](https://www.conventionalcommits.org/):

```
feat: add SMALL, RANK, PERCENTILE, QUARTILE and MODE functions
fix: shift conditional-format ranges once on row/column insert
perf: read sheetData with a raw XmlReader
```

The prefix is what labels the PR, and the label decides both which section of the release
notes the change lands in and how the version is bumped:

| Prefix | Label | Release notes section | Version bump |
|---|---|---|---|
| `feat:` | `enhancement` | New Features | minor |
| `fix:` | `bug` | Bug Fixes | patch |
| `perf:` | `performance` | Performance | patch |
| `docs:` | `documentation` | Documentation | patch |
| `chore:` `build:` `ci:` `style:` `refactor:` `test:` | `chore` | Other Changes | patch |
| any prefix with `!`, e.g. `feat!:` | `breaking` | Breaking Changes | minor (pre-1.0) |

Labelling is automatic — the **PR Autolabel** workflow applies the label from the title when
the PR is opened or edited. You can override it by changing the label by hand; the label
always wins over the title.

Contributors are credited automatically: every entry is attributed to its author, and
first-time contributors get a "New Contributors" section in the release.

## Changelog

User-visible changes belong in `CHANGELOG.md` under `## Unreleased`, in the appropriate
`###` section (`Added`, `Fixed`, `Performance`, `Upgrade Guide`, …). This is written by
hand — it is the place for the detail that a one-line release note can't carry: why the
change matters, migration steps, before/after examples.

Don't add a version heading yourself. The release workflow rolls `## Unreleased` into a
dated version heading when the release is published.

## Releasing

Releases are one click, from the Actions tab:

1. **Publish Release** workflow → *Run workflow*. Leave *version* empty to take the version
   Release Drafter resolved from the merged PRs, or set it explicitly.

   > **The first release must set *version* explicitly.** Release Drafter resolves the previous
   > version from published GitHub Releases, not from git tags. Until one exists it sees none
   > and drafts `v0.0.1`. The workflow refuses any version that isn't an increase over the
   > highest existing tag, so this fails safe — but you have to supply the real version once.
   > After the first release is published, resolution is automatic.
2. Run it once with **dry-run** ticked (the default). It resolves the version and prints the
   changelog roll without changing anything.
3. Before the real run, open the release draft and edit the notes if you want — the draft is
   published as-is, so manual edits survive.
4. Re-run with **dry-run** unticked. The workflow rolls the changelog, commits and tags it on
   `main`, packs the packages, pushes them to NuGet, publishes the release notes, and attaches
   the `.nupkg`/`.snupkg` files.

Pre-release builds go through the **Pre-release** workflow instead (alpha/beta/rc, also with a
dry-run option).

All five packages — `XLibur`, `XLibur.Bundle`, and the three font engines — version in lockstep
from the `v*` tag and are published together. `XLibur.Bundle` references the font engine by
project, so its published dependency is whatever version the font package was packed at; keeping
one version across the set is what guarantees that dependency is a version that actually exists
on NuGet.

## Test Conventions

* Tests use [NUnit](https://nunit.org/) 4.x.
* New tests should use the **constraint model** (`Assert.That(actual, Is.EqualTo(expected))`) rather than the legacy classic asserts (`Assert.AreEqual(expected, actual)`). The actual value goes first, the expectation inside the constraint. Much of the existing suite still uses the classic style; it is being migrated incrementally, so please don't add new classic asserts.

## Setting up the pre-commit hook

The repository includes a pre-commit hook that automatically formats staged C# files with `dotnet format`. After cloning, enable it by running:

```bash
git config core.hooksPath .githooks
```

## Mutation Testing

[Stryker.NET](https://stryker-mutator.io/docs/stryker-net/introduction/) runs daily in CI to measure test effectiveness. To run locally:

```sh
# Restore the Stryker tool (first time only)
dotnet tool restore

# Set StrykerEnabled to disable TreatWarningsAsErrors so mutants can compile
# PowerShell:  $env:StrykerEnabled="true"
# CMD:         set StrykerEnabled=true
# Bash:        export StrykerEnabled=true

# Run mutation testing with the default config
dotnet stryker -f stryker-config.json

# Run against a specific file or folder
dotnet stryker -f stryker-config.json --mutate "XLibur/Excel/Cells/**/*.cs"
```

Reports are generated in `StrykerOutput/` — open the HTML report to see surviving mutants.

## Working with Excel file internals
Excel files (`.xlsx` and `.xlsm`) are zip packages. You can easily verify this by renaming the extension any Excel file to `.zip` and opening the file in your favourite `.zip` file editor.

Internally, the file contains files (also known as parts) that represent different entities in the Excel framework, for example `workbook.xml` and `table1.xml`. The [OpenXML specification](https://www.ecma-international.org/publications/standards/Ecma-376.htm) documents all these parts and their contents.

Making changes to the XLibur code may change the input or output of the package parts. For example if you add support for a currently unsupported element, you will have to ensure that you read the appropriate package part into the XLibur model and also support writing of the package parts to the file.

### Comparing the internals of Excel files

A XLibur developer will often want to compare the internals of 2 similar Excel files. For example if you want to compare the output of a specific package part before and after your code changes. The long, difficult way would be to extract the package parts of the 2 files and manually compare the relevant parts. To ease this, we recommend this tooling stack:

- [Total Commander](https://www.ghisler.com/download.htm)
- [WinMerge](http://winmerge.org/downloads) version `2.14.0`, because subsequent versions for [some reason](https://bitbucket.org/winmerge/winmerge/issues/152/displayxmlfiles-plugin-not-included-with) excludes the required `DisplayXMLFiles.dll` plugin.
- Set Total Commander [to use WinMerge](https://superuser.com/questions/238039/can-i-replace-internal-diff-in-total-commander-with-a-custom-tool) as its compare tool.
- In WinMerge, enable `Plugins > Automatic Prediffer`

Now, to compare 2 similar, but not exact Excel files:

- In Total Commander, navigate to the 1st file in the left-hand pane and the 2nd file in the right-hand pane.
- Press `Ctrl+PageDown` to "enter" the package. You should see, among others, a `[Content_Types].xml` file in both panes.
- You can now compare all package parts by selecting `Commands > Synchronise Dirs...`. Press `Compare`. This will do a full, recursive comparison. You can filter out parts that are identical. 
- You can select an item that differs and press `Ctrl+F3` to open the two parts in WinMerge and see the exact comparison of the part's contents. The XML files should automatically reformat/reindent to ease the comparison instead of showing the entire XML contents on a single line. This is the reason for requiring the `DisplayXMLFiles.dll` plugin.
- In Total Commander, you can also navigate to specific files in the left-hand and right-hand panes and select `File > Compare by Content...`. This will open WinMerge directly.
- Note that since WinMerge reformats the XML, it does so in a temporary file. If you make changes to the contents of any of the 2 panes in WinMerge and save the file, it will not be saved back into the Excel file.

## Reconciling Test Files

XLibur uses a set of [reference .xlsx files](https://github.com/XLibur/XLibur/tree/main/XLibur.Tests/Resource) for comparison for some of the unit tests. Sometimes when you update the XLibur codebase, e.g. a bugfix, the reference test files maybe become obsolete. When running unit tests and the generated file doesn't match the reference file, you will have to update the reference file. You should do this only after inspecting the differences between the generated and reference files in detail and confirming that each change is indeed the expected behaviour. Check the new files visually (e.g. in Excel) and through XML comparison before overwriting the reference files.

