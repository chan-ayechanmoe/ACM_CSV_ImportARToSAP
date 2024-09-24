# ACM_CSV_ImportARToSAP
CSV Import AR To SAP By Aye Chan Moe


This is external window application to read CSV file and create an Item Type AR Invoice in SAP Business One system.

It includes the following features:
(1) Validation if BP not exist, assign a default BP
(2) Validation if Item not exist, assign a default item
(3) CSV file consist of below
	SN, Cust Code, Cust Name, Posting Date ,Item No , Price
	1,BP1,Business Partner1,24/09/2024,Item1,100 
	1,BP1,Business Partner1,24/09/2024,,50 
	2,,No defined BP,24/09/2024,Item1,100
	2,,No defined BP,24/09/2024,Item1,100


