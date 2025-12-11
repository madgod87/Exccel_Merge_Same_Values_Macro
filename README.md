# Excel Merge Same Values Macro

A robust and intelligent VBA macro for Microsoft Excel that automates the task of merging adjacent cells containing identical values. Designed with safety and efficiency in mind, it supports both vertical and horizontal merging and includes an automatic failsafe mechanism.

## 🚀 Key Features

*   **Smart Direction Detection**: Automatically detects whether to merge vertically or horizontally based on the shape of your selection.
    *   *Tall selections* (more rows than columns) trigger **Vertical Merging**.
    *   *Wide selections* (more columns than rows) trigger **Horizontal Merging**.
*   **Batch Processing**:
    *   Select multiple columns to merge duplicates within each column independently.
    *   Select multiple rows to merge duplicates within each row independently.
*   **🛡️ Built-in Failsafe**: **Undo is not possible with macros.** To protect your data, this macro **automatically creates a backup copy** of your active sheet before making any changes. If you don't like the result, simply delete the active sheet and work from the backup.
*   **Performance Optimized**: Disables screen updating and alerts during execution for maximum speed on large datasets.

## 📦 Installation & Setup Guide

### Option 1: Add to a Single Workbook (Easiest)
*Use this if you only need the macro for one specific project.*

1.  **Open your Excel file**.
2.  **Enable the Developer Tab** (if you don't see it):
    *   Right-click anywhere on the top ribbon (the menu with Home, Insert, etc.).
    *   Select **Customize the Ribbon...**
    *   On the right side, check the box for **Developer**.
    *   Click **OK**.
3.  Click the **Developer** tab > **Visual Basic** button (far left).
4.  In the window that pops up, look at the left panel ("Project Explorer"). Right-click on **VBAProject (YourFileName.xlsx)**.
5.  Select **Insert** > **Module**.
6.  Copy the code from the `excelMergeSameValues.vba` file in this repository.
7.  Paste it into the big white text area.
8.  **Save as Macro-Enabled**: When saving your file, go to **File > Save As** and choose **Excel Macro-Enabled Workbook (*.xlsm)**. If you save as a normal Excel file, the code will be deleted!

### Option 2: Make Available in ALL Excel Files (Recommended)
*Use this to have the macro available every time you open Excel, forever.*

Most beginners don't have a "Personal Macro Workbook" yet. Here is how to create one:

1.  **Create the Personal Workbook**:
    *   Open a blank Excel file.
    *   Go to **Developer** tab > **Record Macro**.
    *   In the "Store macro in" dropdown, strictly select **Personal Macro Workbook**.
    *   Click **OK**.
    *   Now, just click any cell, then immediately click **Stop Recording** (top left of Developer tab).
    *   *Congratulations! You just forced Excel to create your hidden "Personal.xlsb" file.*
2.  **Add the Code**:
    *   Press `ALT + F11` to open the Visual Basic editor.
    *   On the left, look for **VBAProject (PERSONAL.XLSB)**.
    *   Expand the folders **Modules** > **Module1** (this is the dummy one you just recorded).
    *   Double-click `Module1`.
    *   Delete the dummy code inside (it will look like `Sub Macro1... End Sub`).
    *   **Paste** the `excelMergeSameValues.vba` code there.
    *   Click the **Save** icon (diskette) in the top toolbar to save your Personal workbook.
    *   Close the VBA editor.

## 📖 How to Use

1.  **Select the cells** you want to fix/merge.
    *   *Tip*: You can select a whole column (click the letter 'A' at the top) or just a block of data.
2.  Run the Macro:
    *   Press `ALT + F8`.
    *   You should see `PERSONAL.XLSB!MergeSameValues` (if you used Option 2) or just `MergeSameValues`.
    *   Double-click it or select it and hit **Run**.

## 🚀 Key Features

*   **Smart Direction Detection**: Automatically detects whether to merge vertically or horizontally based on the shape of your selection.
    *   *Tall selections* (more rows than columns) trigger **Vertical Merging**.
    *   *Wide selections* (more columns than rows) trigger **Horizontal Merging**.
*   **Batch Processing**:
    *   Select multiple columns to merge duplicates within each column independently.
    *   Select multiple rows to merge duplicates within each row independently.
*   **🛡️ Built-in Failsafe**: **Undo is not possible with macros.** To protect your data, this macro **automatically creates a backup copy** of your active sheet before making any changes. If you don't like the result, simply delete the active sheet and work from the backup.
*   **Performance Optimized**: Disables screen updating and alerts during execution for maximum speed on large datasets.

## ⚠️ Important Notes
*   **Backup Naming**: The backup sheet will be named something like `SheetName_Bak_123456`. You can safely delete this if the operation was successful.
*   **Data Loss in Merge**: Standard Excel merging works by keeping the top-left value and discarding the rest. This macro follows that rule but does it automatically for groups of identical values, so no unique data is lost (since the values were identical anyway).

## 📄 License
MIT License - Free to use and modify.
