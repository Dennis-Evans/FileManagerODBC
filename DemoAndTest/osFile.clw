
   member('fmOdbcDemo')

   map
     module('os')
       CreateFile(*cstring lpFileName, long dwDesiredAccess, long dwShareMode, long lpSecurityAttributes, long dwCreationDisposition, long dwFlagsAndAttributes, long hTemplateFile),pascal,raw,handle,name('CreateFileA')
       ReadFile(HANDLE hFile, *cstring lpBuffer, long nNumberOfBytesToRead, *long lpNumberOfBytesRead, long lpOverlapped),pascal,raw,long,name('ReadFile')
       Closehandle(handle f),pascal,long,proc,name('CloseHandle')
       GetLastError(),long,pascal,name('GetLastError')
       getFileSize(HANDLE hFile, *long lpFileSizeHigh),long,pascal,name('GetFileSize')
       WriteFile(HANDLE hFile, *cstring lpBuffer, long nNumberOfBytesToWrite, *long lpNumberOfBytesWritten, long lpOverlapped = 0),pascal,raw,bool,name('WriteFile')
       SetFilePointer(handle hFile, long lDistanceToMove, *long lpDistanceToMove, long dwMoveMethod),pascal,ulong,name('SetFilePointer')
     end
      
   end

CrLf equate('<13><10>')
! ---------------------------------------------------
! reads the file to the end of the file
! ---------------------------------------------------
readFileToEnd  procedure(*cstring fileName, handle f, *cstring sqlCode, long fSize) 

bytesRead  long,auto
errorc     long,auto
retv       bool,auto

  code
  
  retv = ReadFile(f, sqlCode, fSize, bytesRead, 0)
  if (retv = false)
    errorc = GetLastError()
    halt(1, 'unable to read the file, number bytes read ' & bytesRead & ' Error Code ' & errorc & '.')
  end 

  return
! end readFileToEnd -------------------------------------------

! ---------------------------------------------------
! creates the file using the file name input
! if the file exists it is overwritten
! ---------------------------------------------------
makeFile procedure(*cstring filename) ! handle

f handle,auto

  code

  f = CreateFile(fileName, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)  

  return f
! end openSqlFile ------------------------------------------------------

! ---------------------------------------------------
! writes a line of text
! ---------------------------------------------------
Write  procedure(handle f, string text)

textMsg       cstring(MAX_LOG_MESSAGE_LENGTH)  ! arbitray size but large enough
bytesWritten  long,auto  
placeHolder   &long  ! parameter is used for large files and the log file will not be large enough
                     ! so just a place holder

  code 
    
  if (setFilepointer(f, 0, placeHolder, FILE_END) <> INVALID_SET_FILE_POINTER)
    textMsg = text
    if (writeFile(f, textMsg, len(textMsg), bytesWritten, 0) = false) 
      halt(1, 'Unable to write to the log file.')  
    end 
  else 
    halt(1, 'Unable to move the log file pointer.')  
  end  

  return
! ene Write ------------------------------------------

! ---------------------------------------------------
! append crlf to the text and calls write 
! ---------------------------------------------------
WriteLine procedure(handle f, string text)

  code 
  
  write(f, clip(text) & CrLf)

  return
! end WriteLine ---------------------------------------------

! ---------------------------------------------------
! opens the file using the file name input 
! ---------------------------------------------------
openFile procedure(*cstring filename) ! handle

f             handle,auto

  code

  f = CreateFile(fileName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)  

  return f
! end openSqlFile ------------------------------------------------------
 
! ---------------------------------------------------
! reads the file size of the file input. 
! ---------------------------------------------------
findFileSize procedure(handle f) ! long

fileSize  long,auto
sizeX     &long

  code

  fileSize = GetFileSize(f, sizeX)

  return fileSize
! end findFileSize -------------------------------------------------------

! ---------------------------------------------------
! frees the file handle 
! ---------------------------------------------------
CloseFile procedure(handle f) 

  code

  CloseHandle(f)

  return
! end closeFile -------------------------------------------  