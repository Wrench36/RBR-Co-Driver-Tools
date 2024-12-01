# RBR-Co-Driver-Tools
A suite of tools to create and edit co-driver mods for Richard Burns Rally


This is an early pre-alpha state upload of Co-Driver Tools (CDT), meant to demonstrate the progress on the project.

Currently, CDT can:
- Generate an excel file sheet-by-sheet from existing ini's (for organising)
- Generate an exccel file with n sheets from the "package ini file" (the one in pacenotes/config/pacenote/packages)
- Generate ini's from an excel file
- Record co-driver sound pack in python
  - Auto trims and applies radio effect automatically!
- Assist in recording co-driver sound packs
- Rename co-driver sound files automatically
  - When recorded in a DAC and split by note

 A video covering the current features is hosted on youtube here:
https://youtu.be/6VCtJgB5Pxw
TODO:
- Add an option to create a script file from a package ini, all done with one click.
  - This will require selecting any ranges as well. There is no single file which determines all of the ini's in use, so it'll have to be done that way.
    - Look to a later YT video on that. 
- Find a workaround for the end-user having to get rid of the "@" symbols that break the nextID formula
