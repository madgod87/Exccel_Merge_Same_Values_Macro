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

## 📦 Installation

1.  Open your Excel Workbook.
2.  Press `ALT + F11` to open the **Visual Basic for Applications (VBA)** editor.
3.  In the "Project" pane on the left, right-click on your workbook name (e.g., `VBAProject (YourWorkbook.xlsx)`).
4.  Choose **Insert** > **Module**.
5.  Copy the code from `excelMergeSameValues.vba` in this repository.
6.  Paste it into the new module window.
7.  Close the VBA editor.

## 📖 How to Use

### Basic Usage
1.  **Select the range** of cells you want to merge.
2.  Press `ALT + F8` to open the Macro dialog.
3.  Select `MergeSameValues` from the list.
4.  Click **Run**.

### Advanced Scenarios
*   **Merging Multiple Columns**: Select Columns A, B, and C entirely. Run the macro. It will merge identical values in Column A, then Column B, then Column C separately. It will *not* merge a cell from A with a cell from B.
*   **Merging Headers**: Select a range of headers horizontally. If you refer to the same category across multiple sub-columns, the macro will merge them into a single centered header cell.

## ⚠️ Important Notes
*   **Backup Naming**: The backup sheet will be named something like `SheetName_Bak_123456`. You can safely delete this if the operation was successful.
*   **Data Loss in Merge**: Standard Excel merging works by keeping the top-left value and discarding the rest. This macro follows that rule but does it automatically for groups of identical values, so no unique data is lost (since the values were identical anyway).

## 📄 License
MIT License - Free to use and modify.
