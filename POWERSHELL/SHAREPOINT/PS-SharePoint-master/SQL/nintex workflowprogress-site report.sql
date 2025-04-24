select siteid, count (*) from WorkflowProgress WP

inner join WorkflowInstance WI on WI.InstanceID = WP.InstanceID group by siteid order by 2