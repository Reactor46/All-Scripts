$printers = get-wmiobject win32_printer -ComputerName misfrisco1, miseciallen1, mismck1, ecirkwldc, miseciplano1

foreach ( $printer in $printers )
    {
    $driver
