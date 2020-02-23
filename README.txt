
'ΑΛΛΑΓΕΣ ΣΤΟ ΤΙΜ 
'    [ENTITYUID] [varchar](40) NULL,
'	[ENTITYMARK] [varchar](43) NULL,
'	[ENTITY] [int] NULL,
'	[AADEKAU] [float] NULL,
'	[AADEFPA] [float] NULL,
'	[ENTLINEN] [int] NULL,
'	[INCMARK] [nvarchar](43) NULL,


ΑΠΟ ΤΙΜ => INVyyyyddmmhhmm.xml
απαντηση ΑΑΔΕ   INVyyyyddmmhhmm.xml => apantSendInv.xml


 apantSendInv.xml ενημερωση ΤΙΜ (ENTITY , ENTITYMARK )
                  δημιουργία INC.XML (ενημερωση ΤΙΜ (EntLineN) )

απαντηση ΑΑΔΕ   INc.xml => apantIncome.xml


apantIncome.xml  ενημερωση ΤΙΜ ( incMARK ) αντιστοιχιση ( (EntLineN) <=> apantIncome.xml  )
