# Surge
An SRG, STIG, SCAP, and OVAL Import, markup, conversion, and export tool
This creates a XAML gui once executed. There are three buttons along the top of the window. These are "FILE", "MAKE", and "PROFILE / FIX INFO"

The "FILE" menu allows for selecting files of appropriate types for use in the tool. It also allows save and export functionalities.

The "MAKE" menu allows for creation of new STIG spec files. There are also add and remove profile buttons... if unfamiliar, you can think of profiles as groups of checks; Not every check shows in every profile. There are also and, remove, and duplicate rule buttons. The add and remove rule buttons are straightforward; it will add a new check, or remove a selected check from the selected profile. Duplicate rule migrates the selected check into every existing profile. If duplicated, and a user doesn't want the check in all profiles, it must be manually pruned or removed with the remove rule button.

The "PROFILE / FIX INFO" button opens a flyout where the intended profile can be selected. It also holds Check/Rule info, Rule, and Fix Information and display, while leaving the Check/Rule selection tools visible.

The main body of the interface provides entries for asset related info, Rule/check selection, auditor comments entry, and script results, if/where applicable.

Via commandline, the tool allows preselection of a STIG file as the first parameter, such as 
`.\surge.ps1 C:\PATH\TO\STIG\stig.xml`
