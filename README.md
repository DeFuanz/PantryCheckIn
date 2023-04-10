# PantryCheckIn

PantryCheckIn (Pantry visitors) is a client management software program written and developed to help a food pantry local to my area by allow volunteers to easily manage clients and keep track of requirements such as days between visits and whether or not the forms required have been collected.

This project was written and donated at no cost to a local Non-Profit. 

**The Business Problem:**

Previously, the pantry used an excel spreadsheet to manage each client. This included adding, removing, and updating clients information such as ID's and names as well as adding the date they came in and marking a box if their form was collected.
While this worked for a while, the pantry grew to over 12000 clients and statred seeing errors such as duplicates, rows shifting, accidental deletions, and more. This quickly became problomatic as some of these errors would go unoticed and cause conflictions with when users could return, if their forms were collected, and volunteers having to take time to 
comb back through the file to try and fix them.

**The Solution:**

PantryCheckIn was created to help solve all of the collected business problems. Developed in C#, the software solution uses Winforms with Sqlite to enforce business logic such as assurance checks before changes, keeping track of visit dates and only allowing certain actions if criteria is met, as being able to change business days, checking if users already exist when adding a new user, and exporting the data to their original spreadsheet form or as SQL tables if needing to drift away to a new software or revert to excel.

**Image**

Below is the image a volunteer will see when they first load into the software 

![image](https://user-images.githubusercontent.com/76855046/231010338-760fe6e2-21c3-4e27-9d9f-b4511ad02602.png)

**Development**

The solution was developed over a series of roughly 8 months where the first 6 months consisted of developing a working system while the other 2 months were used for testing, obtaining feedback from multiple users and board members, and creating iterations until an agreed upon "definition of done" was reached. 

The solution continues to be used to this day and has successfully helped the non-profit reduce new client visits as well as increase the speed and accuracy of managing current client visits. 

This project was started before taking any advanced programming/engineering classes and so while there were attempts at early understandings of three tier architecture, they are not great versions as project completion and functionality was the main goal as opposed to scalability.

Development had ceased after the definition of done was reached and, while there was a few bugs here and there, greatly satisfied expections for the board members and volunteers alike.
