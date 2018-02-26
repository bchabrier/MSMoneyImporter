# MSMoneyImporter
Automatically import your bank account history into Microsoft Money!

With Microsoft Money sunset, automated download of bank account history was no more supported. Only few banks support the OFX or QIF formats, and often, the files that are produced are not always fully understood by Microsoft Money. This makes the follow-up of your bank account with Microsoft Money quite cumbersome.

`MSMoneyImporter` uses [Boobank](http://weboob.org/applications/boobank) to connect to any supported bank, download the account history in OFX format, adapt the format to Microsoft Money, and imports the file in Microsoft Money.

# Usage
```batch
C:\MSMoneyImporter> msmoneyimporter.bat
Getting the list of bank accounts from boobank...
                                Account                     Balance    Coming
---------------------------------------------------------+----------+----------
             00235648765@cragr CCHQ                         453.43
             04687951322@cragr LDD                           65.21
             64971342165@cragr CEL                          308.99
1345679513298655@societegenerale Livret                       114.38
---------------------------------------------------------+----------+----------
                                             Total (EUR)    942.01       0.00
-------------------
Account 00235648765@cragr (CCHQ)
Creating backup of comptabilité.mny...
Importing C:\Users\bruno\Downloads\CCHQ_00235648765@cragr.ofx into Money (2 transaction(s))...
-------------------
Account 04687951322@cragr (LDD)
Importing C:\Users\bruno\Downloads\LDD_04687951322@cragr.ofx into Money (1 transaction(s))...
-------------------
Account 64971342165@societegenerale (Livret)
Importing C:\Users\bruno\Downloads\Livret_64971342165@societegenerale.ofx into Money (4 transaction(s))...

C:\MSMoneyImporter>
```

# Installation

# Prerequisites
`MSMoneyImporter` uses [Boobank](http://weboob.org/applications/boobank), which in turn require [Python 2](https://www.python.org/downloads/).

# Supported versions of Microsoft Money
`MSMoneyImporter` has been tested with Microsoft Money 2005.

