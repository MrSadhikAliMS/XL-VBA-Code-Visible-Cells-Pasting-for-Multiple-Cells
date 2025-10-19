# XL-VBA-Code-Visible-Cells-Pasting-for-Multiple-Cells
Use this updated VBA code to copy all columns (A:G) from Sheet1 into the visible “Mon” rows on Sheet2, pasting them starting from column B:

Quest:    What is this?
Ans:      Its a XL VBA code to copy all columns (A:G) from Sheet1 into the visible “Mon” rows on Sheet2, pasting them starting from column B:.
          We call it "XL VBA Code Visible Cells Pasting Multiple Cells"

Quest:    Usage of this XL VBA Code?
Ans:      XL VBA CODE, By using this code you can copy data from Sheet1 to only visible (unhidden) cells in Sheet2 where the day is "Mon" using the script.
          Or, Loops through visible “Mon” cells in Sheet2!A:A.
              Copies full row data from Sheet1!A:G sequentially (Day 1 → Day 2 → Day 3...).
              Pastes starting at Sheet2!B, preserving alignment with visible rows.
              This handles all seven columns exactly as the single-column version previously did.

Quest:    How to Use it?
Ans:      In Xl, press Alt+F11(to open MS VBA).
          On top left, Under "File" we can see a Project - VBAProject.
          On that, we can able to find our current Workbook.,
          Right click on the current Workbook->Insert->Module.
          And just paste this MS VBA Code.
          And press F5 for Run the script.
