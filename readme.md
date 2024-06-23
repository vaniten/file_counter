# count_files_by_type.py
Traverses a directory, counting the types of files found and returns a CLI and CSV breakdown of:  count of files, year (last modified), total size of files, total pages (or slides), total rows and total columns (.xls & .xslx). Ouputs:
1. Summary of all file types found, count of individual files, page count, and count of docs with 1-2 pages, 3-5 pages, and >5 pages.
2. Page range table (same as data in summary table).
3. Breakdown of files by year.

## Usage
`python script.py <directory_to_search> [output_file]`

## Dependencies
- datetime
- PyPDF2 
- python-docx 
- python-pptx 
- xlrd 
- openpyxl 
- prettytable 
- pycryptodome

## Sample Output

```
Summary Table:
+-----------+-------+--------------------+-------------+------------+---------------+-----------+-----------+----------+
| File Type | Count | Total Size (bytes) | Total Pages | Total Rows | Total Columns | 1-2 Pages | 3-5 Pages | >5 Pages |
+-----------+-------+--------------------+-------------+------------+---------------+-----------+-----------+----------+
|    pdf    |  1910 |     3237688954     |    11267    |     0      |       0       |    1238   |    321    |   347    |
|    docx   |   68  |      13965128      |     1161    |     0      |       0       |     50    |     5     |    13    |
|    xlsx   |   43  |      2507568       |      0      |   24991    |      172      |     0     |     0     |    0     |
|    xls    |   36  |      2879488       |      0      |   13160    |      1555     |     0     |     0     |    0     |
|    jpg    |   26  |      24641000      |      0      |     0      |       0       |     0     |     0     |    0     |
|    png    |   4   |      7025814       |      0      |     0      |       0       |     0     |     0     |    0     |
|    mp4    |   4   |     769368040      |      0      |     0      |       0       |     0     |     0     |    0     |
|    pptx   |   1   |       348781       |      35     |     0      |       0       |     0     |     0     |    1     |
|    ppt    |   1   |      2485248       |      0      |     0      |       0       |     0     |     0     |    0     |
|    doc    |   1   |       48640        |     114     |     0      |       0       |     0     |     0     |    1     |
|   TOTALS  |  2094 |     4060958661     |    12577    |   38151    |      1727     |    1288   |    326    |   362    |
+-----------+-------+--------------------+-------------+------------+---------------+-----------+-----------+----------+

Page Range Table:
+-----------+-----------+-----------+----------+
| File Type | 1-2 Pages | 3-5 Pages | >5 Pages |
+-----------+-----------+-----------+----------+
|    pdf    |    1238   |    321    |   347    |
|    doc    |     0     |     0     |    1     |
|    docx   |     50    |     5     |    13    |
|    pptx   |     0     |     0     |    1     |
+-----------+-----------+-----------+----------+

Year 2024 Detail:
+-----------+------+-------+--------------------+-------------+------------+---------------+-----------+-----------+----------+
| File Type | Year | Count | Total Size (bytes) | Total Pages | Total Rows | Total Columns | 1-2 Pages | 3-5 Pages | >5 Pages |
+-----------+------+-------+--------------------+-------------+------------+---------------+-----------+-----------+----------+
|    pdf    | 2024 |  164  |      70165995      |     1011    |     0      |       0       |    101    |     36    |    27    |
|    docx   | 2024 |   43  |      2458108       |      54     |     0      |       0       |     39    |     4     |    0     |
|    xlsx   | 2024 |   25  |       619683       |      0      |    302     |       98      |     0     |     0     |    0     |
|    mp4    | 2024 |   4   |     769368040      |      0      |     0      |       0       |     0     |     0     |    0     |
|    doc    | 2024 |   1   |       48640        |     114     |     0      |       0       |     0     |     0     |    1     |
|   TOTALS  | 2024 |  237  |     842660466      |     1179    |    302     |       98      |    140    |     40    |    28    |
+-----------+------+-------+--------------------+-------------+------------+---------------+-----------+-----------+----------+
```