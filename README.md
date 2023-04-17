# files_dirs.vbs
This will list all directories/files under a given directory.

Usage:
cscript files_dirs.vbs (ROOTDIR) (output xls file)

Output file layout
 |　No.　|　column name　|　content　|
 | --- | --- | --- |
 | 0 | DIR/FILE | 'DIR'/'File' |
 | 1 | SEQNO | sequential number |
 | 2 | PARENT_SEQNO | parent sequential number |
 | 3 | DEPTH | depth (starting from ROOT) |
 | 4 | FILE_TYPE | file extention (xlsx, txt, etc) |
 | 5 | SIZE | file size |
 | 6 | LAST_MODIFIED | last modified date |
 | 7 | OWNER | owner of the file |
 | 8 | FILE_NAME | filename |
 | 9 | LINK (file) | Link to the file |
 | 10 | LINK (dir) | Link to the directory |
 | 11 | DIRNAME | directory name |
 | 12 | Path1 | path1 (starting from ROOT) |
 | 13 | Path2 | path2 (starting from ROOT) |
 | 14 | Path3 | path3 (starting from ROOT) |
 | ... | ... | ... | 

# xls2db.vbs
This extracts data from an Excel file, output these data with cell index (row/column index).

Usage:
```
cscript xls2db.vbs (data filename(xlsx)) (output filename(xlsx))
```

## Output Data Layout

 | No. | item | content |
 | --- | --- | --- |
 | 0 | FILENAME | filename |
 | 1 | SHEET_IDX | sheet index |
 | 1 | SHEETNAME | sheetname |
 | 2 | ROW | row index |
 | 3 | COLUMN | column index |
 | 4 | FORMAT | cell format |
 | 5 | VALUE | value |
 
## Sample Input/Output

### Sample Input: 

sheet: "from"

<img width="329" alt="data_from" src="https://user-images.githubusercontent.com/87534698/232361961-ed3ea744-1d89-4f2c-a87c-cba1febe40e6.png">

sheet: "from2"

<img width="284" alt="data_from2" src="https://user-images.githubusercontent.com/87534698/232391814-e63cfcb5-bf81-46f2-a684-e14c7f45dc5f.png">

### Sample Output:


