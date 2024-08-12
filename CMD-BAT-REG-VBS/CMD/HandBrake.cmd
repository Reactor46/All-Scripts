FOR /F "tokens=*" %%G IN ('DIR /B /S *.avi') DO "C:\Program Files\Handbrake\HandBrakeCLI.exe" -i "%%G" -o G:\NAS\XXX-MP4\NEW\"%%G".mp4 --preset="Very Fast 1080p30"
FOR /F "tokens=*" %%G IN ('DIR /B /S *.wmv') DO "C:\Program Files\Handbrake\HandBrakeCLI.exe" -i "%%G" -o G:\NAS\XXX-MP4\NEW\"%%G".mp4 --preset="Very Fast 1080p30"
FOR /F "tokens=*" %%G IN ('DIR /B /S *.mkv') DO "C:\Program Files\Handbrake\HandBrakeCLI.exe" -i "%%G" -o G:\NAS\XXX-MP4\NEW\"%%G".mp4 --preset="Very Fast 1080p30"
FOR /F "tokens=*" %%G IN ('DIR /B /S *.mov') DO "C:\Program Files\Handbrake\HandBrakeCLI.exe" -i "%%G" -o G:\NAS\XXX-MP4\NEW\"%%G".mp4 --preset="Very Fast 1080p30"


for file in `ls /Users/Username/Desktop/Folder`; do $(/Applications/HandBrakeCLI -v -i /Users/Username/Desktop/Folder/${file} -o /Users/audiovideo/Desktop/Converted/"${file%.vob}.mp4" -e x264 -b 3000 -T -E faac -B 192 -R 48 -d slow -5 -8 medium); done
