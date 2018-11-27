# Comparison between Korean address DB and uploaded excel file

We will update Korean address database.
First, the existing table is as follows.

sido(city), gugun(district), dong(street or avenue), updated(for update date) and exist(1-keep or NULL-delete).

> existing..database

<img width="50%" src="https://user-images.githubusercontent.com/39694718/49071923-646f2800-f272-11e8-8dcd-dbac6e84d21c.png">

Also, I created new database and table for updating list. If the exist column is 1 then that row should be deleted and if 0 then that row should be added to existing database.

> newtable..for_address_update

<img width="50%" src="https://user-images.githubusercontent.com/39694718/49072115-e2cbca00-f272-11e8-8ba4-2c4c00c16bed.png">

> address_excel.asp

Insert the changed address data based on the uploaded Excel file into newtable..for_address_update

With the exist attribute, we can delete or insert rows at the address table by the address_updated.asp page(but this page is made by other developer so I would not upload that page) which is shown below.

<img width="100%" alt="image" src="https://user-images.githubusercontent.com/39694718/49074801-cfbbf880-f278-11e8-8b2c-abc43a08f7d8.png">


### The excel file that I used

https://www.mois.go.kr/frt/bbs/type001/commonSelectBoardList.do?bbsId=BBSMSTR_000000000052

Click the latest post and download the jscode[yyyymmdd].zip file.
After unzip the downloaded file, you must use 2 files(KIKcd_B.[yyyymmdd].xlsx 's KIKcd_B sheet and KIKcd_H.[yyyymmdd].xlsx 's KIKcd_H sheet). If the file names or sheet names are changed it should be changed on the code. These files are as follows and I uploaded example files at the Repo.

<img width="100%" src="https://user-images.githubusercontent.com/39694718/49072784-5fab7380-f274-11e8-9a0f-85261a2de5d8.png">

---

## Comparison process
<img width="50%" src="https://user-images.githubusercontent.com/39694718/49073276-671f4c80-f275-11e8-966d-09f095b7c6ce.png">

Upload 2 source files to update original database and then click update button.

<img width="70%" alt="image" src="https://user-images.githubusercontent.com/39694718/49075384-06ded980-f27a-11e8-8cbb-614feb6cbfeb.png">

You can see the progress.
