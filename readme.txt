This is the Visual Basic 5 original project that can generate the DISRowList.dll library.
MSVBVM50.dll is needed in your system directory (on XP already present, on Win7+, google it and just copy to the System32 directory).
The dll is 32bits only (no 64bits VB, damn you Microsoft); on 64 bits systems, you can use the code modules by injecting them directly in your VB project.
Registe DISRowList.dll with the regsvr32 utility (open a command prompt in admin mode, got to the appropriate directory and type "regsvr32 DISRowList.dll").

The VBA2013 subdirectory contains an Access database that has the full 32/64bits VBA modules and testdriver.

Sincerly,
Francesco Foti.