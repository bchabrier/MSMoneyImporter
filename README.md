# MSMoneyImporter
Automatically import your bank account history into Microsoft Money!

With Microsoft Money sunset, automated download of bank account history was no more supported. Only few banks support the OFX or QIF formats, and often, the files that are produced are not always fully understood by Microsoft Money. This makes the follow-up of your bank account with Microsoft Money quite cumbersome.

`MSMoneyImporter` uses [Boobank](http://weboob.org/applications/boobank) to connect to any supported bank, download the account history in OFX format, adapt the format to Microsoft Money, and imports the file in Microsoft Money.

# Usage
```batch
C:\MSMoneyImporter> msmoneyimporter.bat
```

# Installation

# Prerequisites
`MSMoneyImporter` uses [Boobank](http://weboob.org/applications/boobank), which in turn require [Python 2](https://www.python.org/downloads/).

# Supported versions of Microsoft Money
`MSMoneyImporter` has been tested with Microsoft Money 2005.

