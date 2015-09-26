Collection of helpful Excel tools for automating common tasks
=============================================================

###Core features:
- Sorting data ranges that contain merged cells;
- Copying data ranges from one Workbook/Worksheet to another;
- Coloring cells depending on the user's ID;
- Flexible persistent storage of the user's preferences.

###System Requirements:
- Should work on all versions of Microsoft Office Excel.

Main Interface
--------------

![Main Interface Image][MainInterfaceImageId]

Basic Usage
-----------

Let's consider an example that illustrates worksheet sorting depending on the user's id number. First of all, you must fill in all the required settings. The main steps for doing this:
- Select all worksheets you wish to sort on the right panel of the main program window;
- Click on the "Settings" button;
- If you have previously stored some preferences you will see them in the next subwindow;
- Select all the fields you need to change and click on the "Edit" button;
- Save all changes and click "Start sorting".

#####Main Terms Used in This Program
There are a couple of terms, which you can see in the "Edit" panel, that need clarification. These include: Base Cell, Serial Number Cell, Top Left Range Cell, Right Bottom Range Cell. 
The following image can be helpful in understanding their meanings:

![Helper Image][HelperImageId]

**Note:** The project is open source under the [MIT License (MIT)](https://github.com/storix/excel-extra-utils/blob/master/LICENSE).


[MainInterfaceImageId]: http://s14.postimg.org/708824iip/Screenshot_from_2015_09_26_12_07_00.png  "Main Interface"
[HelperImageId]: http://s30.postimg.org/bnflwdbnl/Screenshot_from_2015_09_26_14_14_34.png  "Helper Image"