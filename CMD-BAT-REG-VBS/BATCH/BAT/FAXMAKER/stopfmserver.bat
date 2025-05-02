@echo off
color 1f
SC \\usonvsvrfax01 stop fmserver
SC \\usonvsvrfax01 start fmserver
SC \\usonvsvrfax01 query fmserver
cmd /k
