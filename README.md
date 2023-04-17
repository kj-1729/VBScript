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
