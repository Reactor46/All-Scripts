@echo off
color 1f
SC \\usonvsvrfax01 stop fmmta
SC \\usonvsvrfax01 start fmmta
SC \\usonvsvrfax01 query fmmta
cmd /k
