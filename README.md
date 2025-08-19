The program is called Dayı. Its purpose is simple: Dayı checks whether the programs defined in the code are installed on the computer. It then displays installed and missing programs in separate sections, along with a percentage indicator in the top-right corner. The program also shows system information such as the PC name and Windows version, making it especially useful for bulk device checks in factories.

When exporting to Excel, it writes data in the following structure (this can be customized):

sheet.Cell(1, 1).Value  = "Date";
sheet.Cell(1, 2).Value  = "PC Name";
sheet.Cell(1, 3).Value  = "User Name";
sheet.Cell(1, 4).Value  = "IP Address";
sheet.Cell(1, 5).Value  = "MAC Address";
sheet.Cell(1, 6).Value  = "Installed Programs";
sheet.Cell(1, 7).Value  = "Missing Programs";
sheet.Cell(1, 8).Value  = "Windows Version";
sheet.Cell(1, 9).Value  = "Domain Name";
sheet.Cell(1, 10).Value = "Printer";
sheet.Cell(1, 11).Value = "Driver Status";
sheet.Cell(1, 12).Value = "Windows Update Status";
sheet.Cell(1, 13).Value = "Serial Number";
sheet.Cell(1, 14).Value = "Success Percentage";
sheet.Cell(1, 15).Value = "Old PC";
sheet.Cell(1, 16).Value = "Old Serial Number";


All data will be written to the same Excel file. If the file is already open, a backup Excel will be created, and the results will later be merged — ensuring multiple users can work simultaneously without conflict.
If the same computer is saved again, Dayı updates the existing row instead of creating a new one.

The source code is currently in Turkish, but it can easily be modified to English if preferred.
