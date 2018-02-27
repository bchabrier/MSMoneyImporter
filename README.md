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
`MSMoneyImporter` requires [Boobank](http://weboob.org/applications/boobank), which in turn requires [Python](https://www.python.org/downloads/). 

`MSMoneyImporter` automatically checks that the dependencies are correctly installed, and if not, propose to do the installation. Note that you might have to run `MSMoneyImporter` in elevated mode (i.e. from a command prompt run as administrator).

You can as well do the installation manually as described below. However, the steps to install boobank are not straightforward, so I recommend to let `MSMoneyImporter` do it for you.

Once all dependencies installed, it is necessary to configure the backends by running boobank.


## Python
The current stable version of [Boobank](http://weboob.org/applications/boobank) runs on Python 2.7. Download and install Python 2.7 from this [link](https://www.python.org/downloads/).

## Boobank
[Boobank](http://weboob.org/applications/boobank) can be installed from this [link](http://weboob.org/applications/boobank).


# Prerequisites
`MSMoneyImporter` uses [Boobank](http://weboob.org/applications/boobank), which in turn requires [Python](https://www.python.org/downloads/).

# Supported versions of Microsoft Money
`MSMoneyImporter` has been tested with Microsoft Money 2005.

