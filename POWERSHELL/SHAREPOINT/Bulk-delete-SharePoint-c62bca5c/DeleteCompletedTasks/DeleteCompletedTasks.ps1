######################################################
#Description  : Batch delete all Tasks list items with related workflow history items for all tasks with completed status and before specific date

#Instructions  :      

# 1-Set site url , relative url for task list and relative url for history items list variables in the script paramters area
# 2-Script will delete all items until sepcific date , you can specify this date by set the $DeleteBeforeDate varaible in the script paramters area
# 3-Set the varaible $LogOnly=$true for logging only without delete , this will log only all items to delete and will log the batch messages 
# 4-$DeletedItemsLogPath varaible is contain the log file for all task items and related history items which will delete
# 5-$TaskItemsBatchXmlLogPath varaible is contain the log file for the tasks items batch message that will send to SharePoint 
# 6-$HistoryItemsBatchXmlLogPath varaible is contain the log file for the history items batch message that will send to SharePoint 

#Author       : Wael Mohamed Abdullah (http://waelmohamed.wordpress.com)
#Release Date : 29 / 9 / 2013
######################################################

[void][System.reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

$TodayDate=(Get-Date -Format "dd-MM-yyyy h.mm.ss tt")

###############################################################Script Paramters################################################################################################################
                                                                                                                                                                                              #
$siteUrl="http://www.test.com"                                                        #Site URL                                                                                               #
$tasksListUrl="/Lists/List"                                                           #Tasks Relative List URL                                                                                #
$workflowHistoryListUrl="/WorkflowHistory"                                            #Workflow history list relative url                                                                     #
$DeleteBeforeDate="2013-07-25"                                                        #Delete items before date , format: YYYY/MM/dd                                                          #
$LogOnly=$false                                                                       #if $true the script will write logs only without delete                                                #
$TaskItemsBatchXmlLogPath="C:\test\TaskItemsBatchXml " + $TodayDate +".txt"           #Log path for the tasks items batch file for check                                                      #
$HistoryItemsBatchXmlLogPath="C:\test\HistoryItemsBatchXml " + $TodayDate +".txt"     #Log path for the history items batch file for check                                                    #
$DeletedItemsLogPath="C:\test\DeletedItemsLog "+ $TodayDate +".txt"                   #Log path for the tasks items and related history items that will delete                                #
                                                                                                                                                                                              #
###############################################################################################################################################################################################

#define SharePoint site , web objects and tasks , history lists objects
$site = new-object Microsoft.SharePoint.SPSite($siteUrl) 
$web = $site.rootweb 
$tasksList = $web.GetList($siteUrl+$tasksListUrl)
$tasksListID=$tasksList.ID.ToString()
$workflowHistoryList = $web.GetList($siteUrl+$workflowHistoryListUrl)
$workflowHistoryListID=$workflowHistoryList.ID.ToString()


#query and get list items for tasks list that not completed and before a specific date
$tasksItemsCAML='<Where>
                     <And>
                        <Leq>
                          <FieldRef Name="Created" /><Value IncludeTimeValue="TRUE" Type="DateTime">'+ $DeleteBeforeDate +'T12:00:00Z</Value>
                        </Leq>
                        <Eq>
                           <FieldRef Name="Status" /><Value Type="Choice">Completed</Value>
                        </Eq>
                     </And>
                   </Where>
                   <OrderBy><FieldRef Name="Created" Ascending="TRUE" /></OrderBy>'

$query=new-object Microsoft.SharePoint.SPQuery 
$query.Query=$tasksItemsCAML
$query.ViewAttributes="Scope='Recursive'"
$query.RowLimit=5000

#define to counter varaiables to count actual deleted task items and history items for logging them in the end of the log file
$DeletedTasksItemsCount=0
$DeletedHistoryItemsCount=0


do
{
    $tasksItems=$tasksList.GetItems($query) 
    $query.ListItemCollectionPosition=$tasksItems.ListItemCollectionPosition

    #retrived tasks items count
    $listItemsTotal = $tasksItems.Count

    #define the string builder to build the batch xml for delete tasks items
    $TaskItemsBatchXml=New-Object System.Text.StringBuilder
    $TaskItemsBatchXml.Append("<?xml version=`"1.0`" encoding=`"UTF-8`"?>")
    $TaskItemsBatchXml.Append("<Batch OnError=`"Continue`">")
   
    #start logging
    $msg= "Start at "+  (Get-Date -Format "dd/MM/yyyy h:mm:ss tt") + " calculating "+$listItemsTotal+" task items"
    Add-Content $DeletedItemsLogPath $msg 
    Add-Content $DeletedItemsLogPath "--------------------------------------------------------------------------------------------------"


    #loop through tasks items to collect the batch xml for both tasks and history items
        
        $x = 0
        for($x=0;$x -lt $listItemsTotal; $x++)
        {
         try
         {

            #tasks fields values
            $WorkflowItemId=$tasksItems[$x]["WorkflowItemId"].ToString()
            $taskID=$tasksItems[$x].ID.ToString()
            $taskTitle=$tasksItems[$x].Title.ToString()

            #Build the batch delete xml file for tasks list
            $TaskItemsBatchXml.Append("<Method>")
            $TaskItemsBatchXml.Append([System.String]::Format("<SetList Scope=`"Request`">{0}</SetList>",$tasksListID))
            $TaskItemsBatchXml.Append([System.String]::Format("<SetVar Name=`"ID`">{0}</SetVar>",$taskID))
            $TaskItemsBatchXml.Append("<SetVar Name=`"Cmd`">Delete</SetVar>")
            $TaskItemsBatchXml.Append("</Method>")

            #Write to deleted items log
            $deleteTaskItemMsg ="=>> Task "+ $taskTitle +" with ID " + $taskID + " calculated"
            Add-Content $DeletedItemsLogPath $deleteTaskItemMsg -Encoding UTF8
           
            ####################################################### Delete History Items #########################################################
            
            #query history items that related to the current task item
            $workflowHistoryCAML='<Where><Eq><FieldRef Name="Item" /><Value Type="Integer">'+$WorkflowItemId+'</Value></Eq></Where>'
            $historyQuery=new-object Microsoft.SharePoint.SPQuery 
            $historyQuery.Query=$workflowHistoryCAML

            $historyItemsCount=0
            $historyItems=$workflowHistoryList.GetItems($historyQuery) 
            $historyItemsCount=$historyItems.Count
            
            
            Write-Host  "Task ID "  $taskID " has "  $historyItemsCount "workflow items"

            if ($historyItemsCount -gt 0)
            {
                #define the string builder to build the batch xml for delete history items
                $HistoryItemsBatchXml=New-Object System.Text.StringBuilder
                $HistoryItemsBatchXml.Append("<?xml version=`"1.0`" encoding=`"UTF-8`"?>")
                $HistoryItemsBatchXml.Append("<Batch OnError=`"Continue`">")

                #loop through history items that related to the current task item in the outer loop
                for($i=$historyItemsCount-1;$i -ge 0; $i--)
                {
                    #history items fields values
                    $historyItemID=$historyItems[$i].ID.ToString()

                    try
                    {
                        #Build the batch delete xml file for history list
                        $HistoryItemsBatchXml.Append("<Method>")
                        $HistoryItemsBatchXml.Append([System.String]::Format("<SetList Scope=`"Request`">{0}</SetList>",$workflowHistoryListID))
                        $HistoryItemsBatchXml.Append([System.String]::Format("<SetVar Name=`"ID`">{0}</SetVar>",$historyItemID))
                        $HistoryItemsBatchXml.Append("<SetVar Name=`"Cmd`">Delete</SetVar>")
                        $HistoryItemsBatchXml.Append("</Method>")

                        #Write to deleted items log
                        $deleteWorkflowHistoryMsg ="> History item "+$historyItemID+" for task ID " + $taskID +" calculated"
                        Add-Content $DeletedItemsLogPath $deleteWorkflowHistoryMsg
                        Write-Host  $deleteWorkflowHistoryMsg

                        #count history items that will actually delete
                        $DeletedHistoryItemsCount++
                    } 
                    catch [Exception]
                    {
                        #log the error and contuine
                        $errorMsg="> ERROR: History item "+ $historyItemID +" "+  $Error[0]   
                        Add-Content $DeletedItemsLogPath $errorMsg
                    }
                } #end of history items foreach

                #close the history batch xml tag
                $HistoryItemsBatchXml.Append("</Batch>")

                Add-Content $HistoryItemsBatchXmlLogPath $HistoryItemsBatchXml.ToString()
            
                #check if the script will log only or log and delete 
                if (!$LogOnly) 
                {
                    $web.ProcessBatchData($HistoryItemsBatchXml.ToString())
                }

            ################################################################# End Of Delete History Items ########################################

            #write to log a speartor between each task items 
            Add-Content $DeletedItemsLogPath "--------------------------------------------------------------"
            
            } #end of if condations for history items count > 0
           
             #count tasks items that will actually delete
             $DeletedTasksItemsCount++ 
              
            } #end of try
            catch [Exception]
            {
            #log the error and contuine
            $errorMsg="> ERROR: Task item " + $taskID +" "+ $Error[0] 
            Add-Content $DeletedItemsLogPath $errorMsg
            }
        } #end of task items foreach


        #close the tasks items and history items batch files
        $TaskItemsBatchXml.Append("</Batch>")

        Write-Host "Deleting task items................."

        Add-Content $TaskItemsBatchXmlLogPath $TaskItemsBatchXml.ToString()

        #check if the script will log only or log and delete 
        if (!$LogOnly)
        {
             $web.ProcessBatchData($TaskItemsBatchXml.ToString())
        }

    } while ($query.ListItemCollectionPosition -ne $null)

 
#log the staticts for actual deleted tasks items and history items count
Add-Content $DeletedItemsLogPath "--------------------------------------------------------------------------------------------------"
$msg=  "End of "+$DeletedTasksItemsCount +" task items and "+$DeletedHistoryItemsCount +" history items at "+ (Get-Date -Format "dd/MM/yyyy h:mm:ss tt")
Add-Content $DeletedItemsLogPath $msg 

$web.Dispose() 
$site.Dispose() 