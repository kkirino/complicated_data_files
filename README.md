Parse Complicated Data Files
================
2021-06-15 (update: 2021-06-15)

## Password-encrypted MS Office files

For unlocking passwords of MS Office files, Python library
[`msoffcrypto`](https://pypi.org/project/msoffcrypto-tool/) is useful.

Note that `msoffcrypto` is independent from Java. If you have Java
environment, other tools including `libreoffice` and an R package `xlsx`
can be applied to this task.

``` python
#!/usr/bin/python3
import msoffcrypto

file_path = "./data_file/encrypted.xlsx"
password = "password"

file = msoffcrypto.OfficeFile(open(file_path, "rb"))          # 'read binary'
file.load_key(password = password)
file.decrypt(open("./generated_items/decrypted.xlsx", "wb"))  # 'write binary'
```

Now we can easily read the .xlsx file with Python or R.

``` python
#!/usr/bin/python3
import pandas as pd  # openpyxl is also needed for reading xlsx files 

df = pd.read_excel("./generated_items/decrypted.xlsx")
df.head(6)
```

       id       date type  value
    0   1 2020-12-01    A      1
    1   2 2021-11-24    A      4
    2   3 2019-08-30    B      3
    3   4 2020-03-04    B      5
    4   5 2021-01-12    B      6
    5   6 2020-09-13    C      2

``` r
#!/usr/local/bin/R
library(tidyverse)
library(readxl)

read_xlsx("./generated_items/decrypted.xlsx") %>%
  head()
```

    # A tibble: 6 x 4
         id date                type  value
      <dbl> <dttm>              <chr> <dbl>
    1     1 2020-12-01 00:00:00 A         1
    2     2 2021-11-24 00:00:00 A         4
    3     3 2019-08-30 00:00:00 B         3
    4     4 2020-03-04 00:00:00 B         5
    5     5 2021-01-12 00:00:00 B         6
    6     6 2020-09-13 00:00:00 C         2

## Shift\_JIS-encoded files

We often find Shift\_JIS-encoded data files which were created in
Windows PC. Shift\_JIS encoding is not easy to handle with because
UNIX/Linux-based environment can not read Shift\_JIS encoding.

Using `nkf -g(--guess)` or `file -i` command, we can find the encoding
of the files.

``` bash
#!/bin/bash
nkf -g ./data_file/shis_jis.csv 
```

    Shift_JIS

or

``` bash
#!/bin/bash
file -i ./data_file/shis_jis.csv
```

    ./data_file/shis_jis.csv: application/csv; charset=unknown-8bit

We can change the encoding into UTF-8, which is standard encoding in
UNIX/Linux, with `nkf -w` command in bash.

``` bash
#!/bin/bash
nkf -w ./data_file/shis_jis.csv > ./generated_items/utf_8.csv
```

We confirm the UTF-8 encoding of the file.

``` bash
#!/bin/bash
nkf -g ./generated_items/utf_8.csv 
```

    UTF-8

And we can easily check the contents of the file.

``` bash
#!/bin/bash
head -n 6 ./generated_items/utf_8.csv
```

    13101,100,1000000,トウキヨウト,チヨダク,イカニケイサイガナイバアイ,東京都,千代田区,以下に掲載がない場合
    13101,102,1020072,トウキヨウト,チヨダク,イイダバシ,東京都,千代田区,飯田橋
    13101,102,1020082,トウキヨウト,チヨダク,イチバンチヨウ,東京都,千代田区,一番町
    13101,101,1010032,トウキヨウト,チヨダク,イワモトチヨウ,東京都,千代田区,岩本町
    13101,101,1010047,トウキヨウト,チヨダク,ウチカンダ,東京都,千代田区,内神田
    13101,100,1000011,トウキヨウト,チヨダク,ウチサイワイチヨウ,東京都,千代田区,内幸町

## Large unknown Excel files

Sometimes we get MS Excel files which are large and unknown. Excel files
contain several sheets, but we are not sure how many sheets each excel
file contains.

It is so hard to open xlsx file to check the structures of the files
without high-spec machines.

We can find size of the files using `ls -lh` command. The option `-h`
describe the size as KB or MB.

``` bash
#!/bin/bash
cd data_file
ls -lh large_unknown.xlsx | awk '{print $9 "\t" $5}'
```

    large_unknown.xlsx  1.2G

To confirm the components of the files, Python packages `pandas` and
`openpyxl` are useful.

Although we can get information of the files using R, we can read .xlsx
files with less RAM resources in Python than in R.

``` python
#!/usr/bin/python3
import pandas as pd  # openpyxl is also needed as backend xlsx file handling

path = "./data_file/large_unknown.xlsx"
book = pd.ExcelFile(path) 
sheet_name = book.sheet_names

num_sheet = len(sheet_name)
print("Number of Sheets = ", num_sheet, "\nName of Sheets; ", sheet_name)
```

    Number of Sheets =  8 
    Name of Sheets;  ['Cord_Blood', 'Maternal_Serum', 'Maternal_Urine', 'MT1', 'MT2', 'FT1', 'FFQ', 'Sheet1']

Converting .xlsx to .csv can be achieved whith a Python package
`xlsx2csv`.

``` python
#!/usr/bin/python3
from xlsx2csv import Xlsx2csv
import os

for i in range(num_sheet):
  dest_path = os.path.join("./generated_items/", sheet_name[i] + "_from_xlsx.csv")
  Xlsx2csv(path, outputencoding = "utf-8").convert(dest_path, sheetid = i + 1)
```

Information of the generated .csv files can be seen as below.

``` bash
#!/bin/bash
cd generated_items
ls -lh | grep from_xlsx | awk '{print $9 "\t" $5}'
```

    Cord_Blood_from_xlsx.csv    158K
    FFQ_from_xlsx.csv   453M
    FT1_from_xlsx.csv   84M
    MT1_from_xlsx.csv   123M
    MT2_from_xlsx.csv   130M
    Maternal_Serum_from_xlsx.csv    2.8M
    Maternal_Urine_from_xlsx.csv    1.7M
    Sheet1_from_xlsx.csv    0

Now we can find row and column numbers in each csv file using shell
script, showing Sheet1\_from\_xlsx.csv is a blank file.

``` bash
#!/bin/bash
cd generated_items
for name in *_from_xlsx.csv; do
    echo ${name},`cat ${name} | wc -l`,`head -n 1 ${name} | awk -F ',' '{print NF}'`
done
```

    Cord_Blood_from_xlsx.csv,4829,6
    FFQ_from_xlsx.csv,104063,1964
    FT1_from_xlsx.csv,104063,559
    MT1_from_xlsx.csv,104063,658
    MT2_from_xlsx.csv,104063,691
    Maternal_Serum_from_xlsx.csv,96697,6
    Maternal_Urine_from_xlsx.csv,96856,3
    Sheet1_from_xlsx.csv,0,

## Environment

Finally, we show the information of the environment used for the data
handling here.

    Python 3.8.5
    msoffcrypto-tool 4.12.0 
    openpyxl         3.0.7  
    pandas           1.2.4  
    xlsx2csv         0.7.8  
