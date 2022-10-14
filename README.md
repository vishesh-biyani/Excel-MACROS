# Excel-MACROS

The project uses the powerful feature of MACROS in MS Excel to automate vehicle packing task in addition to creating a DATABASE with each completed task.

## Business Problem
As packing still remains a major task involving workers, it is prone to human errors which can lead to severe losses.
The lack of monitoring in an industrial setting makes it further tough to find the accountable person.

## Explanation
All vehicles need to have 22 components in final packing.
Each component is barcoded.
The barcode is scanned by the packer and only when all items are packed, the ALL OK sticker is printed to send the job to the next station.
At the station, worker has 2 push buttons namely EMERGENCY STOP, Fault Acknowledge to avoid moving back and forth to the PC.

The Push buttons interface with Excel using Arduino UNO Microcontroller and Excel DATA STREAMER add-in.

## MACROS Logic
A password protected sheet is created with a list of all the 22 components.
Whenever a component is scanned, Excel checks whether the component is available in the list and not previously scanned for this vehicle. 
If correctly scanned, Excel appends the scanned component to scanned list and moves ahead 1 row for next component scan.

In case, an item is scanned twice, a DUPLICATE Scanned message is generated which the worker has to acknowledge in order to proceed further.

In case, an item not in the list is scanned, WRONG BOX Scanned message is generated which the worker has to acknowledge in order to proceed further.

Only after all 22 items are scanned properly (not necessarily in order) in additon to VIN Number, 
Excel performs the Total Check, logs the End Packing Time, generates the ALL OK message and appends the entry into the DATABASE.xlsx file.

The user won't see the DATABASE.xlsx file opened as it is opened, data added and then closed, all using MACROS.

## Screenshots
![VBA_1](https://user-images.githubusercontent.com/115135209/195813207-df163ef5-4f6c-447b-b3ed-91acbb76e591.png)

**Main Excel File**


![VBA_2](https://user-images.githubusercontent.com/115135209/195813872-bbcb6159-8e93-4b8f-9c9a-4eb662aabf7e.png)

**DATABASE.xlsx**
