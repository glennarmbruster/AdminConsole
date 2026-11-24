# undo

Pre-change copies of files, one dated folder per change. Restore by copying a
file back over the one in AdminConsole.

NOT an archive. Anything older than 3 days is deleted -- if a change has
survived three days of live use it is not being rolled back, and a folder full
of stale copies is worse than no folder because it invites restoring the wrong
one. Permanent removals go to outputs/obsolete instead.

Prune with:

    find outputs/undo -mindepth 1 -maxdepth 1 -type d -mtime +3 -exec rm -rf {} +
