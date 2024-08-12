exportdir = "c:\vcardsimport"
set fso = createobject("Scripting.FilesystemObject")
set wfile = fso.opentextfile("c:\export.vcf",1,false)
vcfinp = wfile.readall
vcfarry = split(vcfinp,"BEGIN:VCARD",-1,1)
for i = lbound(vcfarry)+1 to ubound(vcfarry)
	Randomize   ' Initialize random-number generator.
	rndval = Int((20000000000 * Rnd) + 1)  
	fname = exportdir & "\" & day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".vcf"
	set nfile = fso.opentextfile(fname,2,true)	
	nfile.writeline("BEGIN:VCARD")
	nfile.write vcfarry(i)
	nfile.close
	set nfile = nothing
next
wscript.echo "done"

