# Developer guidelines

The [OpenXML specification](https://www.ecma-international.org/publications/standards/Ecma-376.htm) is a large and complicated beast. 
In order for XLibur, the wrapper around OpenXML, to support all the features, we rely on community contributions. 

Here are some tips.

* Before starting a large pull request, log an issue and outline the problem and a broad outline of your solution. The maintainers will discuss the issue with you and possibly propose some alternative approaches to align with the XLibur development conventions. 
* Please submit pull requests that are based on the `develop` branch.
* Where possible, pull requests should include unit tests that cover as many uses cases as possible.

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

XLibur uses a set of [reference .xlsx files](https://github.com/XLibur/XLibur/tree/develop/XLibur.Tests/Resource) for comparison for some of the unit tests. Sometimes when you update the XLibur codebase, e.g. a bugfix, the reference test files maybe become obsolete. When running unit tests and the generated file doesn't match the reference file, you will have to update the reference file. You should do this only after inspecting the differences between the generated and reference files in detail and confirming that each change is indeed the expected behaviour. Check the new files visually (e.g. in Excel) and through XML comparison before overwriting the reference files.

