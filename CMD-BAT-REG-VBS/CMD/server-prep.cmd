@echo Off
echo server prep

@call init-path.cmd
@call make-dir.cmd
@call make-optdir.cmd
@call make-share.cmd
@call make-share-acl.cmd
@call copy-tools.cmd

echo *****************************
pause



