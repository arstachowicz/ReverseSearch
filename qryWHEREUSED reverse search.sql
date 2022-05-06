SELECT  tblPARTNUM.ID
       ,tblPARTNUM.partNum
       ,refREV.rev
       ,IIf([obso]=0,"","Obsolete") AS obsol
       ,qryPARTNUM_ALL.partType
       ,qryPARTNUM_ALL.subcategory
       ,qryPARTNUM_ALL.locName
       ,UQ.tbl
FROM qryPARTNUM_ALL
INNER JOIN
(((
	SELECT  tblEP.partNum  AS talonPart
	       ,uSEAL.seal     AS sID
	       ,tblEP.Rev      AS Rev
	       ,tblEP.obsolete AS obso
	       ,"tblEP"        AS tbl
	FROM
	(
		SELECT  id     AS linker
		       ,cseal1 AS seal
		FROM tblEP_CSEAL union
		SELECT  id
		       ,cseal2 AS seal
		FROM tblEP_CSEAL union
		SELECT  id
		       ,cseal3 AS seal
		FROM tblEP_CSEAL
	) AS uSEAL
	INNER JOIN tblEP
	ON uSEAL.linker = tblEP.csealLink
	WHERE (((isNull([seal])) = false)) UNION
	SELECT  tblEP.partNum
	       ,uELEC.elec
	       ,tblEP.Rev
	       ,tblEP.obsolete
	       ,"tblEP"
	FROM
	(
		SELECT  id         AS linker
		       ,electrode1 AS elec
		FROM tblEP_ELECTRODE union
		SELECT  id
		       ,electrode2
		FROM tblEP_ELECTRODE union
		SELECT  id
		       ,electrode3
		FROM tblEP_ELECTRODE union
		SELECT  id
		       ,electrode4
		FROM tblEP_ELECTRODE union
		SELECT  id
		       ,electrode5
		FROM tblEP_ELECTRODE union
		SELECT  id
		       ,groundElectrode
		FROM tblEP_ELECTRODE
	) AS uELEC
	INNER JOIN tblEP
	ON uELEC.linker = tblEP.elecLink
	WHERE (((isNull([elec])) = false)) UNION
	SELECT  tblEP.partNum
	       ,uMISC.idMisc
	       ,tblEP.Rev
	       ,tblEP.obsolete
	       ,"tblEP"
	FROM
	(
		SELECT  id         AS linker
		       ,eleSleeve1 AS idMisc
		FROM tblEP_MISC union
		SELECT  id
		       ,eleSleeve2
		FROM tblEP_MISC union
		SELECT  id
		       ,eleSleeve3
		FROM tblEP_MISC union
		SELECT  id
		       ,eleSleeve4
		FROM tblEP_MISC union
		SELECT  id
		       ,eleSleeve5
		FROM tblEP_MISC union
		SELECT  id
		       ,eleCap1
		FROM tblEP_MISC union
		SELECT  id
		       ,misc1
		FROM tblEP_MISC union
		SELECT  id
		       ,misc2
		FROM tblEP_MISC union
		SELECT  id
		       ,misc3
		FROM tblEP_MISC union
		SELECT  id
		       ,misc4
		FROM tblEP_MISC union
		SELECT  id
		       ,misc5
		FROM tblEP_MISC union
		SELECT  id
		       ,SPACER1
		FROM tblEP_MISC union
		SELECT  id
		       ,spacer2
		FROM tblEP_MISC
	) AS uMISC
	INNER JOIN tblEP
	ON uMISC.linker = tblEP.elecLink
	WHERE (((IsNull([idmisc]))=False)) UNION
	SELECT  tblEP.partNum
	       ,uHalf.partID
	       ,tblEP.Rev
	       ,tblEP.obsolete
	       ,"tblEP"
	FROM
	(
		SELECT  id     AS epID
		       ,capNum AS partID
		FROM tblEP union
		SELECT  id
		       ,basenum
		FROM tblEP
	) AS uHalf
	INNER JOIN tblEP
	ON uHalf.epID = tblEP.ID UNION
	SELECT  tblPRESS.partNum
	       ,uPRESS.idPart
	       ,tblPRESS.rev
	       ,tblPRESS.obsolete
	       ,"tblPRESS"
	FROM
	(
		SELECT  id      AS linker
		       ,partTop AS idPart
		FROM tblPRESS union
		SELECT  id
		       ,partBot
		FROM tblPRESS union
		SELECT  id
		       ,partTOOL1
		FROM tblPRESS union
		SELECT  id
		       ,partTOOL2
		FROM tblPRESS union
		SELECT  id
		       ,partTOOL3
		FROM tblPRESS union
		SELECT  id
		       ,partMISC
		FROM tblPRESS
	) AS uPRESS
	INNER JOIN tblPRESS
	ON uPRESS.linker = tblPRESS.ID
	WHERE (((IsNull([IDPart]))=False)) union
	SELECT  tblLEAK_CHECK.partNum
	       ,uLEAK.idPart
	       ,tblLEAK_CHECK.rev
	       ,tblLEAK_CHECK.obsolete
	       ,"tblLEAK_CHECK"
	FROM
	(
		SELECT  autoid AS linker
		       ,topNum AS idPart
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,bottomNum
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,maniNum1
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,maniNum2
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,miscNum1
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,miscNum2
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,miscNum3
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,miscNum4
		FROM tblLEAK_CHECK union
		SELECT  autoid
		       ,miscNum5
		FROM tblLEAK_CHECK
	) AS uLEAK
	INNER JOIN tblLEAK_CHECK
	ON uLEAK.linker = tblLEAK_CHECK.autoID
	WHERE (((IsNull([IDPart]))=False)) union
	SELECT  tblEP_BATHE.partNum
	       ,uBATH.idPart
	       ,tblEP_BATHE.rev
	       ,tblEP_BATHE.obsolete
	       ,"tblEP_BATHE"
	FROM
	(
		SELECT  ID         AS linker
		       ,assemblyID AS idPart
		FROM tblEP_BATHE union
		SELECT  ID
		       ,base1
		FROM tblEP_BATHE union
		SELECT  ID
		       ,base2
		FROM tblEP_BATHE union
		SELECT  ID
		       ,grndElec1
		FROM tblEP_BATHE union
		SELECT  ID
		       ,grndElec2
		FROM tblEP_BATHE
	) AS uBATH
	INNER JOIN tblEP_BATHE
	ON uBATH.linker = tblEP_BATHE.ID
	WHERE (((IsNull([IDPart]))=False)) ) AS UQ
	LEFT JOIN refREV
	ON UQ.Rev = refREV.ID)
	INNER JOIN tblPARTNUM
	ON UQ.talonPart = tblPARTNUM.ID
)
ON qryPARTNUM_ALL.uniqID = UQ.sID
WHERE (((qryPARTNUM_ALL.linkedID)=[uq]![sID])) OR (((qryPARTNUM_ALL.partNumber)="01-50-0020"))
ORDER BY tblPARTNUM.partNum, refREV.rev; 
