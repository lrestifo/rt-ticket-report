SELECT
	Department,
	SUM(CASE WHEN Prio = 'H' THEN  CreatedThisWeek ELSE 0 END) High_New,
	SUM(CASE WHEN Prio = 'H' THEN     OpenThisWeek ELSE 0 END) High_Open,
	SUM(CASE WHEN Prio = 'H' THEN ResolvedThisWeek ELSE 0 END) High_Closed,
	SUM(CASE WHEN Prio = 'M' THEN  CreatedThisWeek ELSE 0 END) Mid_New,
	SUM(CASE WHEN Prio = 'M' THEN     OpenThisWeek ELSE 0 END) Mid_Open,
	SUM(CASE WHEN Prio = 'M' THEN ResolvedThisWeek ELSE 0 END) Mid_Closed,
	SUM(CASE WHEN Prio = 'L' THEN  CreatedThisWeek ELSE 0 END) Low_New,
	SUM(CASE WHEN Prio = 'L' THEN     OpenThisWeek ELSE 0 END) Low_Open,
	SUM(CASE WHEN Prio = 'L' THEN ResolvedThisWeek ELSE 0 END) Low_Closed,
	SUM( CreatedThisWeek) Total_New,
	SUM(    OpenThisWeek) Total_Open,
	SUM(ResolvedThisWeek) Total_Closed,
	MAX(YEARWEEK(NOW(), 3)) Week_Num
FROM (
  SELECT
    t.Id,
    COALESCE(rr.Organization, 'Unknown') Department,
    CASE
      WHEN t.Priority >= 40 THEN 'H'
      WHEN t.Priority >= 30 THEN 'M'
      ELSE 'L'
    END Prio,
    IF(YEARWEEK(t.Created, 3) = YEARWEEK(NOW(), 3), 1, 0) CreatedThisWeek,
    IF((t.Status = 'resolved' OR t.Status = 'rejected') AND (YEARWEEK(t.Resolved, 3) = YEARWEEK(NOW(), 3)), 1, 0) ResolvedThisWeek,
    CASE
  	  WHEN (YEARWEEK(t.Created, 3) <  YEARWEEK(NOW(), 3)) AND (YEARWEEK(t.Resolved, 3) > YEARWEEK(NOW(), 3)) THEN 1
      WHEN (YEARWEEK(t.Created, 3) <= YEARWEEK(NOW(), 3)) AND (t.Status != 'resolved') AND (t.Status != 'rejected') THEN 1
      ELSE 0
    END OpenThisWeek
  FROM
    Tickets t
    INNER JOIN Users  u ON t.Owner = u.Id
    LEFT OUTER JOIN ObjectCustomFieldValues cf ON t.Id = cf.ObjectId
          AND cf.CustomField = 24
          AND cf.ObjectType = 'RT::Ticket'
          AND cf.Disabled = 0
    LEFT OUTER JOIN (
      SELECT g.Instance TicketId, SUBSTRING_INDEX(GROUP_CONCAT(xm.Organization), ',', 1) Organization
      FROM Groups g INNER JOIN (
        SELECT m.GroupId, m.MemberId, z.Name, CONCAT_WS(' ', z.Country, z.City, z.Address2) Organization
        FROM GroupMembers m LEFT OUTER JOIN Users z ON m.MemberID = z.Id 
      ) xm ON g.id = xm.GroupId
      WHERE g.Domain = 'RT::Ticket-Role' AND g.Type = 'Requestor'
      GROUP BY g.Instance
    ) rr ON t.Id = rr.TicketId
  WHERE
    t.Id = t.EffectiveId
    AND t.Type = 'ticket'
    AND (
      t.Id IN ( 10890, 28036, 30971, 35602, 42262, 44153, 44592, 45573, 46539, 47152,
                47419, 47963, 48982, 49169, 49171, 49172, 49176, 49306, 51493, 52158,
                52469, 52494, 53546, 53702, 53749, 53941, 54009, 54104, 54414, 55153,
                55272, 55636, 56092, 56231, 56268, 56607, 56985, 57253, 59414, 60150,
                60287, 60946, 61211, 61702, 61774, 62081, 62291, 62902, 62992, 63142,
                63181, 63468, 63536, 63575, 64040, 64149, 64225, 64243, 64443, 64447,
                64492, 64639, 66483, 66484, 66485, 66486, 66487, 66488, 66489, 66490,
                66491, 66492, 66493, 66494, 66495, 66496, 66497, 66498, 66499, 66500,
                66501, 66502, 66503, 66504, 66505, 66506, 66507, 66508, 66510, 66511,
                66512, 66513, 66514, 66515, 66516, 66517, 66518, 66519, 66520, 66521,
                66524, 66525, 66526, 66527, 66528, 66529, 66530, 66531, 66532, 66533,
                66534, 66535, 66536, 66537, 66538, 66539, 66540, 66541, 66542, 66543,
                66544, 66545, 66546, 66547, 66548, 66549, 66550, 66551, 66552, 66553,
                66554, 66555, 66556, 66557, 66558, 66718 )
      OR (
        t.Created > '2014-03-16'
        AND t.Id NOT IN (
                10890, 28036, 30971, 35602, 42262, 44153, 44592, 45573, 46539, 47152,
                47419, 47963, 48982, 49169, 49171, 49172, 49176, 49306, 51493, 52158,
                52469, 52494, 53546, 53702, 53749, 53941, 54009, 54104, 54414, 55153,
                55272, 55636, 56092, 56231, 56268, 56607, 56985, 57253, 59414, 60150,
                60287, 60946, 61211, 61702, 61774, 62081, 62291, 62902, 62992, 63142,
                63181, 63468, 63536, 63575, 64040, 64149, 64225, 64243, 64443, 64447,
                64492, 64639, 66483, 66484, 66485, 66486, 66487, 66488, 66489, 66490,
                66491, 66492, 66493, 66494, 66495, 66496, 66497, 66498, 66499, 66500,
                66501, 66502, 66503, 66504, 66505, 66506, 66507, 66508, 66510, 66511,
                66512, 66513, 66514, 66515, 66516, 66517, 66518, 66519, 66520, 66521,
                66524, 66525, 66526, 66527, 66528, 66529, 66530, 66531, 66532, 66533,
                66534, 66535, 66536, 66537, 66538, 66539, 66540, 66541, 66542, 66543,
                66544, 66545, 66546, 66547, 66548, 66549, 66550, 66551, 66552, 66553,
                66554, 66555, 66556, 66557, 66558, 66718 )
        AND (
          t.Queue IN ( 26, 27, 28, 29, 30, 31, 32, 33, 35, 82 )            
          OR (
            t.Queue NOT IN ( 26, 27, 28, 29, 30, 31, 32, 33, 35, 53, 75, 81, 82 )
            AND COALESCE(cf.Content, 'Unknown') IN (
                  'ERP_SAP_Authorization', 'ERP_SAP_EDI_PBS', 'ERP_SAP_EPD',
                  'ERP_SAP_FICO', 'ERP_SAP_Interface', 'ERP_SAP_MM',
                  'ERP_SAP_PP', 'ERP_SAP_Program', 'ERP_SAP_SD', 'ERP_SAP_WM' )
          )
        )
      )
    )
  ) Tkts
WHERE CreatedThisWeek = 1
   OR ResolvedThisWeek = 1
   OR OpenThisWeek = 1
GROUP BY Department
ORDER BY Department
