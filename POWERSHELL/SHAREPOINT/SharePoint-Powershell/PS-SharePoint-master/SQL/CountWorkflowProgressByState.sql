-- Get the number of WorkflowProgress records in DB based on State of WorkflowInstance
-- 2 = 'Running'; 4 = 'Completed'; 8 = 'Cancelled'; 64 = 'Error'

SELECT count(WP.WorkflowProgressID) as Error
  FROM [dbo].[WorkflowInstance] as WFInst
  JOIN [dbo].[WorkflowProgress] as WP ON WFInst.InstanceID = WP.InstanceID
  WHERE WFInst.[State] = 64

SELECT count(WP.WorkflowProgressID) as Cancelled
  FROM [dbo].[WorkflowInstance] as WFInst
  JOIN [dbo].[WorkflowProgress] as WP ON WFInst.InstanceID = WP.InstanceID
  WHERE WFInst.[State] = 8

  SELECT count(WP.WorkflowProgressID) as Completed
  FROM [dbo].[WorkflowInstance] as WFInst
  JOIN [dbo].[WorkflowProgress] as WP ON WFInst.InstanceID = WP.InstanceID
  WHERE WFInst.[State] = 4