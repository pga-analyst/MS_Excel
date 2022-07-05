# MS_Excel
Identify Best 10 Locations to Store Received Product

This file takes modified data from the SCMS Delivery History dataset located on Kaggle and uses made up item locations and sizes. The formulas uses index matches as well as mid to reference specific text in order to return every nth location along with it’s cubing.

Formulas:

•	=IF(AND($Q2="CS",$R2>5),INDEX(Locations,SMALL(IF((Shelf="A")+(Shelf="B"),IF((MID(Locations,1,3)="BD6")+(MID(Locations,1,3)="BD2")+(MID(Locations,1,3)="BD3")+(MID(Locations,1,3)="BD4")+(MID(Locations,1,3)="BD5"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1)),T$1)),IF(AND($Q2="CS",$R2<5,$R2>1),INDEX(Locations,SMALL(IF((Shelf="B")+(Shelf="C"),IF((MID(Locations,1,3)="BD6")+(MID(Locations,1,3)="BD2")+(MID(Locations,1,3)="BD3")+(MID(Locations,1,3)="BD4")+(MID(Locations,1,3)="BD5"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1)),T$1)),IF(AND($Q2="CS",$R2<1),INDEX(Locations,SMALL(IF((Shelf="C")+(Shelf="D"),IF((MID(Locations,1,3)="BD6")+(MID(Locations,1,3)="BD2")+(MID(Locations,1,3)="BD3")+(MID(Locations,1,3)="BD4")+(MID(Locations,1,3)="BD5"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1)),T$1)),IF($Q2="<>",INDEX(Locations,SMALL(IF((MID(Locations,1,3)="BD1"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1),T$1)),""))))

Auto-Batcher

This file takes in the first 248 items from the SCMS Delivery History dataset located on Kaggle as orders. Regardless of the number of items requested, it breaks the order down to be picked by a batch size that does not exceed 3000 cases per batch. Up to 7 batches can be picked based on the case size and item size. This file uses multiple vlookups but the main formula(s) use index matching to group the items as well as arrays containing the items, their case quantities and their totals. The formulas change slightly per batch.

Formulas:

•	=Q4:OFFSET(Q4,CELL("contents",U1)-1,1,1)

•	=CONCATENATE("Case Count: ",SUM(T4:T1000))

•	=IFERROR(INDEX($A$4:$N$1000,SMALL(IF((INDEX(MID($A$4:$N$1000,1,2),,11)<>"NN")*(INDEX($A$4:$N$1000,,13)>=100)*(INDEX($A$4:$N$1000,,13)<>""),MATCH(ROW($A$4:$N$1000),ROW($A$4:$N$1000)),""),ROWS($A$4:A4)),COLUMNS($A$1:A1)),"")

•	=IF(INDIRECT(CELL("address",OFFSET(INDEX(S:S,MAX((S:S<>"")*(ROW(S:S)))),1,-2))):OFFSET(INDIRECT(CELL("address",OFFSET(INDEX(S:S,MAX((S:S<>"")*(ROW(S:S)))),1,-2))),CELL("contents",X1)-1,1,1)<>LOOKUP(2,1/(ISNUMBER(S4:S1000)),V4:V1000),INDIRECT(CELL("address",OFFSET(INDEX(S:S,MAX((S:S<>"")*(ROW(S:S)))),1,-2))):OFFSET(INDIRECT(CELL("address",OFFSET(INDEX(S:S,MAX((S:S<>"")*(ROW(S:S)))),1,-2))),CELL("contents",X1)-1,1,1),"Duplicate")


Amazon Webpage Query

This sample is a basic Web query for Amazon’s Top 30 OTC items. It automatically queries the webpage organizes the query in a readable list that contains the item rating and price. It can be extended to Top 100 if necessary. Vlookups are used but the main function to organize the list is the offset function. Offset is used to return the cell details in a list where the formula references the next 5th row down. (Currently had to adjust the formula in item 7 since data was missing in the query solely for this item).

Formulas:

•	=Query!A17

•	=OFFSET(Query!A$17,(ROW()*0)+$B6,0)

Query Connection String:

•	https://www.amazon.com/Best-Sellers-Health-Personal-Care-OTC-Medications-Treatments/zgbs/hpc/10079990011


