# TimeTracker
Provides script for tracking duration for which you are active on computer. This script can be automated to include in Scheduler which runs it every time user locks or unlocks the system. It generates excel report - 1 for each month. Refer to the readme file for details of how to integrate this script to automatically execute on lock/unlock of system.

# Creating scheduled time tracking
You can create a scheduled task that will run when your computer is locked/unlocked:
	1. Start > Administrative Tools > Task Scheduler
	2. left pane: select Task Scheduler Library
	3. right pane: click Create Task... (NOTE: this is the only way to get the correct trigger)
	4. in the Create Task dialog:
		○ General tab -- provide a name for your task
		○ Triggers tab -- click New... and select On workstation lock/unlock
		○ Action tab -- click New... and click Browse... to locate your script (time_tracker.py)
    ○ Conditions tab -- uncheck Start the task only if the computer is on AC power
    
Create 1 task for workstation lock and 1 for workstation unlock. For both these tasks, time_tracker.py shuould be specified as the script name in Actions tab.
