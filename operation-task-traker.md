# TaskTrackerForm 

```bash
txtTaskName
txtNotes | Notes for task (multiline = True, scrollbars = 2 (vertical)).
lblTotalTime | Display total time spent. Set Caption = "Total Time Spent: 0 hours".
lblRowIndex | Hidden. Stores row index. Set Visible = False.
txtActiveTask | Read-only or hidden. Shows resumed task name. Set Visible = False.
btnStartTask | 
btnEndTask
btnReset
```

## Adding Notes

```plaintext
Adding notes
txtNotes:

MultiLine = True

EnterKeyBehaviour = True

ScrollBars = fmScrollBarsVertical

txtActiveTask:

Locked = True if visible

BackColor light gray (to show it's non-editable)
```

```bash
+--------------------------+
| [ txtTaskName        ]   |
| [ txtNotes           ]   |
| [ lblTotalTime       ]   |
| [ btnStartTask ]          [ btnEndTask ]    [ btnReset ] |
| [ txtActiveTask ]         (hidden or read-only)           |
| [ lblRowIndex   ]         (hidden)                        |
+--------------------------+
```

