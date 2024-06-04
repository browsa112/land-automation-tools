# land-automation-tools
This is currently in development and hasn't had all the bugs worked out yet. I don't see any errors so far yet when testing for a maxmimum of six different counties, but there might be bugs that I missed. I reccomend not using this your only way to generate a mail log, but as a checker to compare if what you did by hand was valid.

If the excel file has a missing lot size, it will be skipped. This might make it so that the on market/sale or sold count is off by a couple of properties (this varies by state and location). I checked one case that was missing a lot size in the csv file, but showed up in the under 1 acre range online and noticed in the description that there were actually two lots in this one sale. I think that was why the lot size was empty in the csv lot size for that listing. 

# Steps 

## 1. Download the excel files sold (I normally select one year ago for sold, but any timeframe can be used) and on market data 
### They should appear in the downloads folder with a similar file name to the below files
<img width="772" alt="Pasted Graphic 1" src="https://github.com/browsa112/land-automation-tools/assets/168380980/ec981b4d-8c85-4010-8f8c-77ba04ea901a">



## 2.  Change the file names so that they match if they were an (on market/sale) or sold downloaded file with the naming standard: county_state_sold or county_state_sale (underscores not dashes) all lowercase
### If the county name is two words then insert a dash between each name in the county ex. San Juan County becomes san-juan_co_sold or san-juan_co_sale

Sold and on market/sale data for Stevens Washington  

### Before:
<img width="772" alt="Pasted Graphic" src="https://github.com/browsa112/land-automation-tools/assets/168380980/c48404e3-792c-4030-a0b3-989cb16d2a31">


### After:
<img width="783" alt="Pasted Graphic 2" src="https://github.com/browsa112/land-automation-tools/assets/168380980/607c1f5e-8af2-4726-bd93-a8b43705575a">


## 3. Put all the excel files into the same folder as the main.py folder (the folder name can be anything, I just used demo because it was the first name that came to mind)
Ex.
<img width="745" alt="baltimore_md_sale csv" src="https://github.com/browsa112/land-automation-tools/assets/168380980/61d22393-d3cd-492a-bec2-dcbfe8241b43">



## 4. Open the terminal (I am using a Mac in the below screenshot)
<img width="579" alt="Terminal" src="https://github.com/browsa112/land-automation-tools/assets/168380980/426d5cd3-16eb-4a53-ae05-5eff48104c75">


### 5. Open the terminal and go to the folder that contain the main.py file (I used cd to get there with the below command on the terminal, but your command will probably be different based on the folder name and location)
<img width="221" alt="Pasted Graphic 5" src="https://github.com/browsa112/land-automation-tools/assets/168380980/ad4045aa-0ae5-47bf-9b93-c521d7e7f795">



### 6. Type “python3 main.py” on the terminal and press enter
<img width="179" alt="Pasted Graphic 6" src="https://github.com/browsa112/land-automation-tools/assets/168380980/141a24f6-8866-4fd3-b925-00327b19e0f1">



### 7. This should print out data on each county 
￼<img width="515" alt="Pasted Graphic 7" src="https://github.com/browsa112/land-automation-tools/assets/168380980/f6cbafc5-ce64-41bc-a3bd-8757b55b3838">



### 8. A new file name “new.xlsx” should appear with the data entered for each county
<img width="756" alt="May 27, 2024 at 939 PM" src="https://github.com/browsa112/land-automation-tools/assets/168380980/7671a3e0-c2a2-4822-b954-fe1724ae82f5">
