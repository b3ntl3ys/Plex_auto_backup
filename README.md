# Plex_auto_backup
 a Python script that allows for an easy-to-use graphical user interface (GUI) to zip files from a Plex media server and export registry keys
it closes the plex media server then backups the registry keys for plex and zips the plex media server folder when done it starts plex media server again

registry is: 
HKEY_CURRENT_USER\Software\Plex, Inc.\Plex Media Server

Plex Media Server:
AppData\Local\Plex Media Server


#convert .py to exe:
pip install pyinstaller

pyinstaller --onefile -w 'plex_auto_backup.py'


The code is provided under an open source MIT license, which grants you the freedom to use, modify, and distribute it according to the terms of the license.
