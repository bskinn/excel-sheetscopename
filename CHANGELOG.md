## CHANGELOG: Sheet-Scoped Name Generator Excel VBA Add-In

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).


### [Unreleased]

...

### [0.1.0] - 2019-01-03

*Initial release*

#### Features
 * Generates sheet-scoped names for all selected cells based on the values
   of their left-hand neighbors

#### Limitations
 * Only supports un-accented ASCII letters

#### Known Bugs
 * RTE if column A is selected
 * RTE if a left-neighbor cell is empty
 * RTE if attempting to name a merged cell
 * RTE if attempting to draw a name from a merged cell

#### Deprecated Capability
 * Allows name generation from a selection with more than one column,
   permitting behavior not anticipated to be particularly useful,
   and probably more confusing that it's worth

#### Internals
 * Automatically converts left-neighbor contents to a valid name
   (in most cases)