# count_files_by_type.py
Traverses a directory, counting the types of files found and returns a CLI and CSV breakdown of:  count of files, year (last modified), total size of files, total pages (or slides), total rows and total columns (.xls & .xslx). Breakdown is provided by year and also a summary of the files in the directory.

## Usage
`python script.py <directory_to_search> [output_file]`

## Dependencies
- PyPDF2 
- python-docx 
- python-pptx 
- xlrd 
- openpyxl 
- prettytable 
- pycryptodome

## Sample Output

```
Year: 2020
+-----------+------+-------+--------------------+-------------+------------+---------------+
| File Type | Year | Count | Total Size (bytes) | Total Pages | Total Rows | Total Columns |
+-----------+------+-------+--------------------+-------------+------------+---------------+
|    pdf    | 2020 |   2   |       188286       |      8      |     0      |       0       |
|   TOTALS  | 2020 |   2   |       188286       |      8      |     0      |       0       |
+-----------+------+-------+--------------------+-------------+------------+---------------+

Summary Totals
+-----------+-------+--------------------+-------------+------------+---------------+
| File Type | Count | Total Size (bytes) | Total Pages | Total Rows | Total Columns |
+-----------+-------+--------------------+-------------+------------+---------------+
|    pdf    |  1910 |     3237688954     |    11267    |     0      |       0       |
|    docx   |   68  |      13965128      |     1161    |     0      |       0       |
|    xlsx   |   43  |      2507568       |      0      |   24991    |      172      |
|    mp4    |   4   |     769368040      |      0      |     0      |       0       |
|    doc    |   1   |       48640        |     114     |     0      |       0       |
|    jpg    |   26  |      24641000      |      0      |     0      |       0       |
|    png    |   4   |      7025814       |      0      |     0      |       0       |
|    xls    |   36  |      2879488       |      0      |   13160    |      1555     |
|    pptx   |   1   |       348781       |      35     |     0      |       0       |
|    ppt    |   1   |      2485248       |      0      |     0      |       0       |
|   TOTALS  |  2094 |     4060958661     |    12577    |   38151    |      1727     |
+-----------+-------+--------------------+-------------+------------+---------------+
```