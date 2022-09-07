# Po Excel Tool (PET)

I created this command line utility to speed up my own translating process using Excel xlsx file.

It can be basically used for:
  - adding multiple [PO](https://www.drupal.org/community/contributor-guide/reference-information/localize-drupal-org/working-with-offline/po-and-pot-files#s-po-files) files to one sheet in Excel file
  - extract required locale(s) from xlsx sheet to [PO](https://www.drupal.org/community/contributor-guide/reference-information/localize-drupal-org/working-with-offline/po-and-pot-files#s-po-files) file(s)

As I often use [Babel](https://babel.pocoo.org/en/latest/index.html) for [extracting](https://babel.pocoo.org/en/latest/cmdline.html#extract) traslated strings which generates [POT](https://www.drupal.org/community/contributor-guide/reference-information/localize-drupal-org/working-with-offline/po-and-pot-files#s-pot-files) file it is possible:
  - to merge [POT](https://www.drupal.org/community/contributor-guide/reference-information/localize-drupal-org/working-with-offline/po-and-pot-files#s-pot-files) file with older [PO](https://www.drupal.org/community/contributor-guide/reference-information/localize-drupal-org/working-with-offline/po-and-pot-files#s-po-files) files before creating spreadsheet table

Workflow is as minimal as possible, you need only two commands:

  - ```pet toxls``` and  ```pet fromxls``` 
  - it uses **all po** files in **actual directory** and **default name** for Excel table is **messages.xlsx**

Ouput table is hopefully :slightly_smiling_face: beautifully formatted and columns´ length is adjusted:

![image](https://user-images.githubusercontent.com/3373705/180677755-a3aeabec-fe66-49e7-b895-3fc927e2d601.png)

## Using

```
pet toxls -h
Usage: pet toxls [OPTIONS] CATALOG

  Convert .PO files to an XLSX file.

  guessing locale for PO files:
      1. "Language" key in the PO metadata,
      2. filename.
  manual locale specification:
      pet toxls cs=basedir/czech/mydomain.po

  pet toxls en.po Bulgarian.po
  - add files en.po and bg.po as en and bg locale to messages.xlsx

  pet toxls en=British.po cs=Czech.po
  - add files British.po and Czech.po as en and cs locale to messages.xlsx

  pet toxls
  - add all po files from current dir to messages.xlsx

Options:
  -c, --comments [translator|extracted|reference|all]
                                  Comments to include in the spreadsheet
  -o, --output FILENAME           Output file  [default: messages.xlsx]
  -m, --msgmerge                  flag for update(merge) from pot file
  -h, --help                      Show this message and exit.
```

```
pet fromxls -h
Usage: pet fromxls [OPTIONS] [LOCALE]...

   Convert a XLS(X) file to a .PO file

    pet fromxls en cs
    - create en.po and cs.po files from messages.xlsx

    pet fromxls en=British.po cs=Czech.po
    - create British.po and Czech.po from message.xlsx

   pet fromxls
   - extract all locales from messages.xlsx to appropriate po files
     in current directory


Options:
  -is, --ignoresheet TEXT  Ignore sheets with specific names.
  -od, --outdir DIRECTORY  output directory for po file  [default: .]
  -if, --inputfile PATH    input xlsx file  [default: messages.xlsx]
  -h, --help               Show this message and exit.
```

## Instalation
```
pip install git+https://github.com/zbyna/po-excel-tool.git#egg=poexceltool
```

As a base for this project was https://github.com/wichert/po-xls used. Thank you!
