$dirName = "D:\Solutions"
foreach ($solution in Get-SPSolution)
{
    $id = $Solution.SolutionID
    $title = $Solution.Name
    $filename = $Solution.SolutionFile.Name
    $solution.SolutionFile.SaveAs("$dirName\$filename")
}


#Read more: https://www.sharepointdiary.com/2011/10/extract-download-wsp-files-from-installed-solutions.html#ixzz7jbcPeMnY