Identify Best 10 Locations to Store Received Product

This file takes modified data from the SCMS Delivery History dataset located on Kaggle and uses made up item locations and sizes. The formulas uses index matches as well as mid to reference specific text in order to return every nth location along with it’s cubing.

Formulas:

•	=IF(AND($Q2="CS",$R2>5),INDEX(Locations,SMALL(IF((Shelf="A")+(Shelf="B"),IF((MID(Locations,1,3)="BD6")+(MID(Locations,1,3)="BD2")+(MID(Locations,1,3)="BD3")+(MID(Locations,1,3)="BD4")+(MID(Locations,1,3)="BD5"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1)),T$1)),IF(AND($Q2="CS",$R2<5,$R2>1),INDEX(Locations,SMALL(IF((Shelf="B")+(Shelf="C"),IF((MID(Locations,1,3)="BD6")+(MID(Locations,1,3)="BD2")+(MID(Locations,1,3)="BD3")+(MID(Locations,1,3)="BD4")+(MID(Locations,1,3)="BD5"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1)),T$1)),IF(AND($Q2="CS",$R2<1),INDEX(Locations,SMALL(IF((Shelf="C")+(Shelf="D"),IF((MID(Locations,1,3)="BD6")+(MID(Locations,1,3)="BD2")+(MID(Locations,1,3)="BD3")+(MID(Locations,1,3)="BD4")+(MID(Locations,1,3)="BD5"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1)),T$1)),IF($Q2="<>",INDEX(Locations,SMALL(IF((MID(Locations,1,3)="BD1"),ROW(Locations)-ROW(INDEX(Locations,1,1))+1),T$1)),""))))
