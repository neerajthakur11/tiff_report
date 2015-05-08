::clean
rmdir /S /Q tiff_report
rmdir /S /Q build
del tiff_report.exe
::make
python setup.py py2exe
move dist tiff_report
"C:\Program Files\WinRAR\Rar.exe" a -sfx -z"xfs.conf" tiff_report tiff_report/
