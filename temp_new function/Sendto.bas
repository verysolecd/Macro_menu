This interface allows to use 'Send To' functionalities through an API.

Example: Set Send=CATIA.CreateSendTo() 
This interface requires the installation of CATIA - PPR xPDM Gateway 1 Product (PX1) or the installation of the CATIA-SmarTeam plugin. In case one of these products is not granted, the first invocation to one of CATIASendToService methods will fail. 
Method Index 
AddFile  Adds a file to the list of the files 'to be copied'.  
GetLastSendToMethodError  Retreives the diagnosis related to the last call to SendToService interface.  
GetListOfDependantFile  Retreives the complete list of the files recursively pointed by the file given in argument to SetInitialFile method.  
GetListOfToBeCopiedFiles  Retreives the complete list of the files that will be copied.  
KeepDirectory  Controls the directory tree structure in the target directory.  
RemoveFile  Removes a file from the list of the files that will be copied.  
Run  Executes the copy action, according to previously set files and options.  
SetDirectoryFile  Positions the destination directory.  
SetDirectoryOneFile  Allows positioning the destination directory for one given file to be copied.  
SetInitialFile  Sets the initial file to be copied.  
SetRenameFile  Renames one file to be copied.  


Methods


Sub AddFile(CATBSTR iPath) 
Adds a file to the list of the files 'to be copied'. This method verifies that the given input file is valid (exists and is not a directory), it recursively adds pointed files. 
Parameters: 
iPath 
: The path of the file to be added to the list of the 'to be copied' files. 
Example: 
Send.AddFile(iPath) 
Sub GetLastSendToMethodError(CATBSTR oErrorParam,long oErrorCode) 
Retreives the diagnosis related to the last call to SendToService interface. 
Parameters: 
oErrorParam 
A parameter string given together with the error code. 
oErrorCode 
The last executed method error code: 
code diagnosis oErrorParam value 
0  action successfully performed :-)  
1  PX1 license not granted   
2  internal error   
5  file already in the list  file name  
6  file is not in the list  file name  
7  empty file list   
8  missing target directory   
9  no common root directory   
10  file does not exist  file name  
11  input is a directory  directory name  
12  directory check failed  directory name  
13  invalid file name  given name  
14  file has no read permission  given name  
36  allocation failed :-(   

Sub GetListOfDependantFile(CATSafeArrayVariant oDependant) 
Retreives the complete list of the files recursively pointed by the file given in argument to SetInitialFile method. Notice : in case AddFile has also been invoked, the files recursively pointed by the added file also are retreived. 
Parameters: 
oDependant 
: The table of dependant files. 
Example: 
Send.GetListOfDependantFile(oDependant) 
Sub GetListOfToBeCopiedFiles(CATSafeArrayVariant oWillBeCopied) 
Retreives the complete list of the files that will be copied. This list matches the list of dependant files, but without the files for which RemoveFile has been invoked. 
Parameters: 
oWillBeCopied 
: The table of the files that will be copied. 
Example: 
Send.GetListOfToBeCopiedFiles(oWillBeCopied) 
Sub KeepDirectory(boolean iKeep) 
Controls the directory tree structure in the target directory. 
Parameters: 
iKeep 
=1: to preserve the relative tree structure of the files.
This option will be effective only if there is a common root directory for all files. 
iKeep 
=0: to copy the files directly in the destination directory 
Example: 
Send.KeepDirectory(ikeep) 
Sub RemoveFile(CATBSTR iFile) 
Removes a file from the list of the files that will be copied. 
Parameters: 
iFile 
: The File (With extension) to be removed from the list of the 'to be copied' files. 
Example: 
Send.RemoveFile(iFile) 
Sub Run() 
Executes the copy action, according to previously set files and options. 
A "report.txt" report file is generated in the specified destination directory. 
Sub SetDirectoryFile(CATBSTR iDirectory) 
Positions the destination directory. This method verifies that the given directory exists. Be careful, if SetDirectoryOneFile method has been previously called, its action is overriden by this SetDirectoryFile call. 
Parameters: 
iDirectory 
: The destination directory where the files will be copied. 
Example: 
Send.SetDirectoryFile(iDirectory) 
Sub SetDirectoryOneFile(CATBSTR iFile,CATBSTR iDirectory) 
Allows positioning the destination directory for one given file to be copied. The file will be copied in the specified target directory. Be careful that using this method implies that the 'KeepDirectory' variable will be automatically set to 0. 
Parameters: 
iFile 
: The name (Name With extension) of the given file. 
iDirectory 
: The directory where this file will be copied. 
Example: 
Send.SetDirectoryOneFile(iFile, iDirectory) 
Sub SetInitialFile(CATBSTR iPath) 
Sets the initial file to be copied. This method verifies that the given input file is valid (exists and is not a directory) 
It generates a complete list of the recursively dependent files to be copied. 
Example: 
This example positions the file of path ipath in the list of 'to be copied' files. All its dependant files will also be added in the list of 'to be copied' files. 
Parameters: 
iPath 
: Full path of the file to be copied. 
 Send.SetInitialFile(iPath)
 Sub SetRenameFile(CATBSTR iOldname,CATBSTR iNewName) 
Renames one file to be copied. The new name may not have invalid characters 
Parameters: 
iOldname 
: The old file name (With extension). 
iNewName 
: The new file name (Without extension). 
Example: 
Send.SetRenameFile(iOldname, iNewName) 
