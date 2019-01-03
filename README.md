# Sheet-Scoped Name Generator -- Excel VBA Add-In

*Lightweight utility to automatically create sheet-scoped Names for selected cells based on their left-hand neighbors.*

The built-in <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F3</kbd> keyboard shortcut, which invokes the
`Formulas > Defined Names > Create from Selection` command, always creates names at the
workbook-global scope.  To the best of this author's knowledge, there is no built-in functionality
for automatically creating *worksheet*-scoped names. This add-in attempts to rectify that omission.

To use, select the cells for which names are to be created and press
<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>N</kbd>.  The name applied to each cell will be
created from the value of the cell to its immediate left.

The binary `.xlam` file for each release can be found on the GitHub page for that release.

Copyright (c) Brian Skinn 2019

License: The MIT License  
See [`LICENSE.txt`](https://github.com/bskinn/excel-sheetscopename/blob/master/LICENSE.txt) for full license terms.

*Sheet-Scoped Name Generator is third-party software, and is neither affiliated with, nor authorized,
sponsored, or approved by, Microsoft Corporation.*
