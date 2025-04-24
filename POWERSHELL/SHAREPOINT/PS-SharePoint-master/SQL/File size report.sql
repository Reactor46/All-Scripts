select physical_name, size,
    CONVERT(DECIMAL(10,2),(size * 8.00) / 1024.00) As UsedSpaceMB
from master.sys.master_files
order by physical_name
