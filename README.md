# grasp-tools-vb6
A Visual Basic 6 IDE Add-In

## Command List

* Tools Menu
    * Find forms
    * Open StartUp Object (If Present)
    * Open MDI Form (If Present)

* Window Menu
    * Close All Code Windows Except the Active One
    * Open Containing Folder of Active CodePane

* Project Window Form Folder
    * Open Containing Folder

* Project Window Module/Class Folder
    * Open Containing Folder

## Limitations

* "Open Containing Folder" functionality assumes that
    * the file name and the project item name are equal
    * all files used in the project reside in the same folder as the project file
    * if it does not find the file, it will insrtuct Explorer to open the project folder.
    
## Remarks

"Find forms" command opens a form in which you can look for project forms by item name or by caption.