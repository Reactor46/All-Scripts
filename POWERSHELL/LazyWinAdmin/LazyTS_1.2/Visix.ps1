 
     
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="MyForm" 
    $MyForm.Size = New-Object System.Drawing.Size(300,300) 
     
 
        $mVisix Machines = New-Object System.Windows.Forms.ListBox 
                $mVisix Machines.Text="ListBox1" 
                $mVisix Machines.Top="30" 
                $mVisix Machines.Left="5" 
                $mVisix Machines.Anchor="Left,Top" 
        $mVisix Machines.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mVisix Machines) 
        $MyForm.ShowDialog()
