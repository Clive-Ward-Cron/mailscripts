# mailscripts

Requires Python 3.8 to run the prep files.
Python versions above 3.8 will **NOT** run on Windows 7.

prep-aga.py & prep-pga.py

Bash Format: `prep-[aga | pga] [aga list filename] -F [alt file name]`

`-F` flag is optional for if the base csv name is different than the default.

---

Requires Python 2.7.17 to run dbf-to-xls file.
The following libraries only run in Python 2:

- xlwt
- xlrd
- xlutils
- dbfpy
