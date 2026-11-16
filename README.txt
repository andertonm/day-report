DAY REPORT - NEW VERSION (Python)            2900 Shore Drive
==============================================================

WHAT IT DOES
Every morning, on its own, the office PC:
  1. picks up last night's Passport report files (MSM, MCM, ISM) from the POS,
  2. keeps a copy of every file in the archive folder,
  3. fills in that day's "Day N" sheet in  C:\Users\Mark\Documents\AEI Data\day_reports\day report <Month> <Year>.xlsm,
     the same way the old dayreport.exe did. The Day sheet and the month file follow the
     date in the Passport report's <EndDate>. That is the date the data applies to
     (Passport closes the day around 21:00-22:00). Example: a report ending Oct 1 goes to
     Day 1 of the October workbook.
  4. writes a short log:  C:\Users\Mark\Documents\AEI Data\day_reports\logs\dayreport.log
It does NOT open Excel and does not need anyone logged in.
HOW THE FUEL SHEETS AND THE SUMMARY GET THEIR NUMBERS
The program writes the raw Passport data into the "Day N" sheet (columns L and T), exactly
like the old dayreport.exe did. The workbook's own formulas then carry those numbers to
the other sheets. The old program never typed into those sheets either:
  * Tank sheets, column B "Sales", row = day + 7 (Day 1 = row 8 ... Day 31 = row 38):
      IR_Regular   = Regular gallons + 67% of Plus     ('Day N'!B6 + 0.67*'Day N'!B7)
      IR_Supreme   = Supreme gallons + 33% of Plus     ('Day N'!B8 + 0.33*'Day N'!B7)
      Tank 3 and 4 = same as IR_Supreme (manifold)
      Diesel       = Diesel gallons                    ('Day N'!B9)
      IR_Supreme2  = 0 (unused tank)
  * Summary: average daily gallons by grade (B4:B8), store/lottery/beer/cigarette totals,
    garage (FLAC) sales, gas purchased to date, all summed or averaged across Day 1 to Day 31.
  * The fuel order button on Summary (calc_gas macro) still works the same way.
These numbers appear when the workbook is opened in Excel, because the program marks the file
"recalculate everything on open". Phone or web previews that do not recalculate may show
old numbers. Open it in Excel.
The program never writes your typed entries: stick inches, deliveries, water, cost/gal,
cash on hand, EFT, lottery serials. Those cells stay exactly as you left them.
After each day it also checks that the tank sheets' Sales cell for that day points at that
day's sheet, and writes a WARNING in the log if it does not (see "KNOWN TEMPLATE ISSUES").

YOUR DAILY ROUTINE (nothing new)
  1. Open this month's day report workbook in Excel as usual.
     (Excel recalculates everything when it opens. Give it a second.)
  2. Go to the Day sheet for the report's date (Day N = the date the report ends on).
     A2 should say "Business date: MMDDYY" with that same date.
  3. Enter cash on hand, EFT, lottery serials and tank stick readings as before.
  4. SAVE AND CLOSE the workbook when you are done. If it is left open overnight,
     the morning run cannot save. It will log an error and try again on the next run.

IF SOMETHING LOOKS WRONG
  * Open C:\Users\Mark\Documents\AEI Data\day_reports\logs\dayreport.log and scroll to the bottom. Every run starts
    with "==== dayreport start" and ends with "exit code N".  0 = fine.
      exit code 1 = a day could not be processed (the reason is on the ERROR line)
      exit code 3 = the POS (192.168.87.11) could not be reached
      exit code 4 = another run was still going
      exit code 5 = one or more PDF reports did not print (they stay in booutbox)
      exit code 6 = the 11:55 PM tank grab could not save tank readings (see "tank grab" lines)
  * To run it by hand: double-click  C:\DayReport\run_dayreport.bat
  * "Day N says MMDDYY but holds different numbers" = that sheet already holds data for
    that date from a different Passport file. It is left alone. Check it, and rerun with --force if the new file is right.
  * Item-level sales (dailysales.dbf) are only written if dailysales_dbf is set in dayreport.ini.
    It is empty by default, so that step is skipped.
  * A day is never filled twice. To redo a day on purpose (for example after Passport
    re-sent a file), put its MSM and MCM files back into
    C:\shell save\passport\booutbox  and run, in a Command Prompt:
        C:\DayReport\run_dayreport.bat --no-fetch --force
  * Before every change, a backup of the workbook goes into  C:\Users\Mark\Documents\AEI Data\day_reports\backups
  * A new month's workbook is created from  C:\Users\Mark\Documents\AEI Data\day_reports\day report template.xlsm
    the first time a file for that month arrives. An existing month file is never replaced.

ONE-TIME SETUP ON "SDS Office PC" (Mark)
  1. Install Python 3.12 (python.org, 64-bit). Tick "Install for all users" and
     "Add to PATH". It should land in  C:\Program Files\Python312\
  2. Make a folder C:\DayReport and copy into it: dayreport.py, run_dayreport.bat,
     requirements.txt, dayreport.ini.sample, DayReport_task.xml
  3. Command Prompt:
        cd C:\DayReport
        "C:\Program Files\Python312\python.exe" -m pip install -r requirements.txt
        copy dayreport.ini.sample dayreport.ini
     Edit dayreport.ini if any folder is different on this PC.
  4. Store the POS login (logged on as the same Windows user the task will run as):
        cmdkey /add:192.168.87.11 /user:BackOffice /pass
     Type the BackOffice password when asked. Then check:  dir \\192.168.87.11\XMLGateway\booutbox
  5. Test without changing anything:
        run_dayreport.bat --dry-run
     then a real run:  run_dayreport.bat   and look at the log.
  6. Create the scheduled task (it runs 05:30 every day; change the time in Task Scheduler if you want):
        schtasks /Create /TN "Day Report" /XML C:\DayReport\DayReport_task.xml /RU "%COMPUTERNAME%\%USERNAME%" /RP *
     (Enter your Windows password. That is what "Run whether user is logged on or not" needs.)
     Or, without the XML:
        schtasks /Create /TN "Day Report" /TR "C:\DayReport\run_dayreport.bat" /SC DAILY /ST 05:30 /RU "%COMPUTERNAME%\%USERNAME%" /RP * /RL LIMITED /F
     Test it now:  schtasks /Run /TN "Day Report"   then read the log.
  7. Disable or delete the old scheduled job / shortcut that runs the old batch file and dayreport.exe,
     so the two versions don't both grab the files.
  If you ever change your Windows password, update the task (Task Scheduler > Day Report >
  Properties > OK, and re-enter the password).

TEXT MESSAGE (SMS) ALERTS - Twilio
  Status: built in, but NO recipients are set, so no texts are sent. Recipients will be added
  after the stakeholder questionnaire.
  Alert types:
    error            run failed (exit code 1)
    pos_unreachable  POS 192.168.87.11 could not be reached (exit code 3)
    collision        a Day sheet already holds different data for that date (not overwritten)
    missing_day      a date is missing from the sequence, or no Passport files were found
    already_running  another run was still active (exit code 4)
    success          optional daily "all good" text with gallons and sales
    missed_run       sent by the WATCHDOG task when this morning's run did not finish OK
  The same alert is not re-sent within dedup_hours (12 by default). Alert types listed in
  quiet_hours_apply_to (success, missing_day) are held during quiet_hours (22:00-07:00).
  They go out after that, on the next run or watchdog. Texts never change the run's result.
  If sending fails, it is only logged.

  Setup (once, logged on as the SAME Windows user the scheduled tasks run as):
  1. Twilio account: buy a number. US texting to cell phones needs the number REGISTERED:
     A2P 10DLC brand + campaign registration for a local number, or toll-free verification for a
     toll-free number. Unregistered numbers are filtered or blocked by the carriers. Registration can
     take days to a few weeks. Start it early. Texts to a number not yet verified may silently not arrive.
  2. cd C:\DayReport
     "C:\Program Files\Python312\python.exe" -m pip install -r requirements.txt
     "C:\Program Files\Python312\python.exe" dayreport.py --setup-twilio
     (asks for Account SID, Auth Token, and the From number. It stores them in Windows Credential
     Manager, never in a file. Environment variables TWILIO_ACCOUNT_SID / TWILIO_AUTH_TOKEN /
     TWILIO_FROM also work as a fallback.)
  3. Test:  "C:\Program Files\Python312\python.exe" dayreport.py --test-sms +1XXXXXXXXXX
  4. Add recipients to dayreport.ini, one [recipient.NAME] section each (see dayreport.ini.sample):
     number in +1XXXXXXXXXX form, and alerts = the types that person should get.
  5. Register the watchdog task. It runs at 07:30, 2 hours after the main 05:30 run, and texts
     "missed_run" if the main run didn't run or didn't finish OK:
        schtasks /Create /TN "Day Report Watchdog" /XML C:\DayReport\DayReport_watchdog_task.xml /RU "%COMPUTERNAME%\%USERNAME%" /RP *
     or:  schtasks /Create /TN "Day Report Watchdog" /TR "C:\DayReport\run_dayreport.bat --watchdog" /SC DAILY /ST 07:30 /RU "%COMPUTERNAME%\%USERNAME%" /RP * /RL LIMITED /F
     If you change the main task's time, keep the watchdog about 2 hours after it.
  Note: if the PC itself is off or has no internet, nothing can text. The watchdog only covers
  "the PC is on but the job didn't run or failed".

KNOWN TEMPLATE ISSUES (in the workbook's own formulas, not the program)
  * IR_Supreme, column B, days 1,2,4,5,6,7 use 33% of Plus from Day 3, and days 16 to 22 use
    Day 15's Plus, instead of the same day's. Supreme tank sales are off on those days.
    Fix in the template: IR_Supreme!B8 should be ='Day 1'!$B$8+0.33*'Day 1'!$B$7 and so on.
  * Diesel!B38 (day 31) is a typed 0 instead of ='Day 31'!$B$9.


CARRYING FORWARD GAS INVENTORY (OPENING INVENTORY, ROW 7)
On every run for a month, if last month's workbook is in the same folder, the program
puts last month's closing gallons into this month's opening cell E7 on IR_Regular, Diesel,
IR_Supreme and IR_Supreme2, and puts IR_Supreme + IR_Supreme2 into "Tank 3 and 4" E7.
- Closing = the stick inches you typed in column D on last month's last calendar day
  (row 37 for a 30-day month, row 38 for a 31-day month), converted to gallons with that
  workbook's IR_Parameters chart, the same way column E does it. The value is written as a
  plain number, not a formula.
- If that last-day reading isn't entered yet, the log says so, and the program fills E7
  on a later run once the reading is there.
- E7 is filled only while it is blank or still holds the template's value. If you type
  your own opening figure, the program leaves it alone.
- A backup is made in backups\ before the workbook is changed, and each carry is logged.


PRINTING THE STAFF PDF REPORTS (added Nov 2026)
-----------------------------------------------
Every run (5:30 AM) also prints every .pdf file in \\192.168.87.11\XMLGateway\booutbox
to the printer "Lexmark in store", then deletes it from booutbox.
  - A copy of each PDF is saved first in  <archive_dir>\printed\YYYY-MM-DD\
  - A PDF is deleted from booutbox ONLY after the printer accepted it. If printing
    fails, the file stays in booutbox, the log has a "PDF print FAILED" line, the run
    ends with exit code 5 and an error alert is sent. It is retried at the next run.
  - --dry-run only lists what it would print. --no-fetch skips printing.
  - Turn off with  print_pdfs = no  in dayreport.ini. Printer name: pdf_printer.
  - This step is separate from the Day Report: one failing does not stop the other.

Printing uses SumatraPDF (free, open source, no Adobe needed, no windows pop up).
SumatraPDF.exe (portable 64-bit, version 3.5.2) is included in this zip - just keep it in
C:\DayReport next to dayreport.py. If it is ever missing, download it again:
  1. Go to https://www.sumatrapdfreader.org/download-free-pdf-viewer
  2. Download "SumatraPDF-3.5.2 64-bit portable" (a .zip), or directly:
     https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64.zip
  3. Unzip it and rename the .exe to SumatraPDF.exe, put it in C:\DayReport.
Test it by hand once (prints one page, no window should open):
  C:\DayReport\SumatraPDF.exe -print-to "Lexmark in store" -silent "C:\some\test.pdf"


TANK WATCHER STICK READINGS (updated Nov 11 2026 - no Gmail, no Google Drive)
-------------------------------------------------------------------------------
Every night at 11:55 PM a second scheduled task, "Day Report Tanks", reads the
public Tank Watcher page https://gastanks.watch/home and saves the 4 tanks'
readings into one small file:
   C:\AEI Data\day_reports\tank_snapshots\tank-snapshot-YYYY-MM-DD.csv
The 5:30 AM Day Report run then writes each tank's height (inches) into
column D of that day's row on IR_Regular, Diesel, IR_Supreme and IR_Supreme2 -
the same cells you used to type by hand. Excel's formulas turn inches into
gallons. "Tank 3 and 4" is left alone; it still sums the two Premium tanks.

  * A cell is filled ONLY when it is blank. A typed value is never overwritten.
    A workbook backup is made before any change.
  * The 11:55 PM grab never opens a workbook and never touches the POS.
  * Heights on the site are whole inches.
  * The time in the file is the time Tank Watcher took the reading.
    A reading between midnight and 5:59 AM counts as the previous day's closing.
  * If the site can't be reached, it retries for about 5 minutes. If it
    still fails, the log has a "tank grab FAILED" line and the task ends
    with exit code 6. If the site is up but its numbers are more than 3 hours
    old (Raspberry Pi 192.168.0.28 stopped sending), that also fails with exit code 6.
  * If last night's file is missing at 5:30 AM, the log has a WARNING and those
    stick cells stay blank for staff. The rest of the run is not affected.
  * The old Gmail / Apps Script / Google Drive route is no longer needed. If you
    installed the Apps Script, you can delete that project at script.google.com.

SETUP (Mark) - about 5 minutes, logged on as Day_Report
  1. Copy the new  dayreport.py  and  DayReport_tanks_task.xml  into  C:\DayReport
     (replace the old dayreport.py).
  2. Open  C:\DayReport\dayreport.ini  in Notepad. Find the tank_fill and
     tank_csv_dir lines (or add them at the end of the [dayreport] part) so they say:
        tank_fill = yes
        tank_csv_dir = C:\AEI Data\day_reports\tank_snapshots
     Save.
  3. Test it. Open Command Prompt and run this one line:
        C:\DayReport\run_dayreport.bat --grab-tanks
     It should end with "tank grab end, exit code 0". Then check the file is there:
        dir "C:\AEI Data\day_reports\tank_snapshots"
  4. Create the nightly task (asks for the Day_Report Windows password):
        schtasks /Create /TN "Day Report Tanks" /XML C:\DayReport\DayReport_tanks_task.xml /RU "%COMPUTERNAME%\Day_Report" /RP *
  5. Optional test of the task itself:
        schtasks /Run /TN "Day Report Tanks"
     then look at the bottom of the log file for "tank snapshot written".
