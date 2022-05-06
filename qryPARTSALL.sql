SELECT  subQry.uniqID
       ,subQry.autoID
       ,subQry.tbl
       ,subQry.partNumber
       ,subQry.partUseage
       ,tblPART_TYPES.partType
       ,tblPART_STORAGE.locName
       ,subQry.linkedID
       ,subQry.notes
       ,subQry.location
       ,refREV.rev
       ,tblSUBCATEGORY.category
       ,subQry.subcategory
FROM
((((
	SELECT  ('C' & autoID) AS uniqID
	       ,autoID
	       ,rev
	       ,'tblPART_156'  AS tbl
	       ,partNumber
	       ,partUseage
	       ,location
	       ,linkedID
	       ,notes
	       ,subcategory
	FROM tblPART_156
	UNION ALL
	SELECT  ('A' & autoID) AS uniqID
	       ,autoID
	       ,rev
	       ,'tblPART_01'   AS tbl
	       ,partNumber
	       ,partUseage
	       ,location
	       ,linkedID
	       ,notes
	       ,subcategory
	FROM tblPART_01
	UNION ALL
	SELECT  ('B'&autoID) AS uniqID
	       ,autoID
	       ,rev
	       ,'tblPART_07' AS tbl
	       ,partNumber
	       ,partUseage
	       ,location
	       ,linkedID
	       ,notes
	       ,subcategory
	FROM tblPART_07) AS subQry
	LEFT JOIN tblPART_TYPES
	ON subQry.partUseage = tblPART_TYPES.ID)
	LEFT JOIN tblPART_STORAGE
	ON subQry.location = tblPART_STORAGE.ID)
	LEFT JOIN refREV
	ON subQry.rev = refREV.ID
)
LEFT JOIN tblSUBCATEGORY
ON subQry.subcategory = tblSUBCATEGORY.ID;
