RT_reporter is a tool for tracking time and automated reporting to RT via the REST interface.
It is composed of two main functional parts, an .xlsx file (a template is part of the repository) for inputting the time spent on individual tasks, and the RT_reporter.py script for the reporting of the input time.

It was created as a personal tool to work with the Cosylab RT server, but feel free to connect to a different RT server by modifying the RT_SERVER constant in the RT_reporter.py (the script should work with any RT server with REST 1.0 interface)

# Inputting time with the .xlsx file

## Defining new projects and tasks
A new project is defined by creating a new *Project* sheet (the template contains three project sheets, *PROJ1*, *PROJ2* and *None*). This is best done by duplicating and renaming an existing *Project* sheet, as to keep the excel formulas. The project name is auto-populated from the name of the *Project* sheet.

To create new tasks, fill out the *Task*, *T_spent*, *T_todo* and the *RT link* columns. The *T_spent* and *T_todo* are formulas, so either copy them from an existing task or use `CTRL+D` to duplicate the formulas from the row above.

As the time is input into the *Time* sheet (see below), the time for the particular tasks is summed in the *Project* sheets. The following fields are reported to the user:
- *Project*: Project name, as defined by the name of the *Project* sheet.
- *TOTAL time input*: Total time, spent on all of the project tasks, in minutes.
- *T_spent*: Total time, spent on a particular project task, in minutes.
- *T_todo*: Difference between *T_spent* and *T_input*, representing the amount of time that still needs to be reported to RT.

The *T_input* field can be filled out manually, to specify the time that was already input into RT. The RT_reporter.py script automatically updates these values based on the time reported when it is run.

## Inputting time, spent on tasks
The time, spent on particular tasks is input on the *Time* sheet. This is done by filling out the following columns:
- *Date*: If starting time tracking for a new day, input a date.
- *Project* and *Task* name as defined in the *Project* sheets.
- *Start* and *Finish* time of the task.
`dw` can be input into the *Project* column if the time should not be recorded as working time.

As the tasks are input the following cells are populated using formulas:
- *\[Worked\]*: The total time spent working on a particular day.
- *\[Minutes\]*: Time spent working on a task.
- *\[Status\]*: Not filled when a project is not defined, *GREEN* if OK, *RED* when the project is not defined, *YELLOW* when the task cannot be found in the specified project, *GRAY* when non-working time is input.

*Note: If the row formulas are corrupted, or new rows need to be added, this is done most easily by selecting a complete row and using `CTRL+D` to duplicate the row above. Note that hidden columns are used in B-E for data processing.*

## Time analysis
The analysis sheet can be used for displaying some basic data on the input time:
- *No. of working days*: Based on the dates input into the *Date* column.
- *No. of working hours*: *No. of working days* \* 8 hours.
- *Total time \[hh:mmh\]*: Total time, input into the *Time* sheet.
- *Difference*: *No. of working hours* - *Total time*, red if negative (less time was input than there were working hours).
- *Average work time*

# Reporting time using RT_reporter.py
The script is run by using `python RT_reporter.py location/of/work.xlsx`
The location of the .xlsx file represents the relative filepath, so for a work.xlsx file, stored in the same folder as the python script, use `python RT_reporter.py work.xlsx`
*Note: openpyxl is used as a library, make sure that it is installed on your system before first use by using `pip install openpyxl` or as appropriate for your system.

After inputting the username and password, the RT_reporter script loops through all of the projects and tasks and reports time to the defined RT tickets. The following information is displayed to the user for each task:
- *Project* and *Task* names
- *RT#* and *RT ticket name*, as defined in RT
- *Time to input*
The script ignores the tasks for which tickets are not defined.
The user must confirm the reporting of time by inputting `c`, with an optional text to be submitted with the reply in the form of `c This is some optional text.`. Any other input will mean the script will skip the reporting for this task.

Once the time for all the tasks is reported, the script updates the *T_input* fields for all the tasks for which the time was reported into RT.

The following errors are handled by the python script, most being self-explanatory or including guidance to resolve:
- *Usage: python RT_reporter.py location/of/work.xlsx*: Displayed at inappropriate usage of the script (inappropriate number of arguments).
- *The .xlsx file cannot be found at the specified location.*
- *Cannot open the .xlsx file, close before proceeding.*
- *Credentials error*
- *Error! Status code received: {HTTP status code}*: The most typical case being 403 (Forbidden), in case VPN is not connected.
- *Invalid ticket number for task '{task}' in RT link: {rt_link}.*
- *Cannot find a valid ticket with number {rt_link} for task '{task}', skipping.*