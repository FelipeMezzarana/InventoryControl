# inventory_control
VBA program to control a factory's warehouse inventory

1. Introduction

This program offers an interface for the user to control and visualize a database with detailed information about a factory's inventory, offering the possibility of simultaneous use by several users.

Its main functions are:
- Add and remove inventory items
- Register new items
- Edit existing Items
- Get information about where the item is stored
- Control missing and/or low quantity items (purchasing planning)
- Generate identification tags
- Many validation tools to prevent user errors

Obs: The program was originally developed in portuguese, and later everything has been translated and adapted to english. Therefore, the Database in this repository (which contains fictional items and serves as an example only) have items named in Portuguese. To start using the program, you must delete all the contents of the database, keeping only its structure. Then, register the desired items.


2 - Requirements
- Office package installed (Access installation is not required)
- Font installation “code 39” for creating barcodes (can be found in a .zip file in this repository)
- The folder containing the database “DATABASE” (found in this repository) must be located in the same path as the main file “inventory _control_v1.56”

3 - User instructions

3.1 Interface overview

![image](https://user-images.githubusercontent.com/90487618/164818163-38e1bbff-f735-4f7d-bf28-3e7b7d490b5b.png)

- 1 - The program was developed for 17'' square screens, which means that this interface may not adapt perfectly depending on the screen used. That's why we have zoom in, zoom out and fullscreen buttons to make the program adjustable on any screen.

- 2 - Open a Useform for inventory control, management and editing.

- 3 - Open a Userform for registering new items

- 4 – unprotects the workbook for editing through a password, so changes can be made to the design, registration of new passwords, change in the names of the database and its folder, etc. (admin password “E1234”)



3.2 Inventory Receipt\Issue

![image](https://user-images.githubusercontent.com/90487618/164836462-92f15a6a-1656-4099-8aed-df03af8dd6e1.png)


This is the Main Userform of the program. Here you can view, edit, add or remove items.

3.2.1 Searching for items:

The three buttons and the three comboboxes offer different ways of filtering and selecting items, we will explain each one individually:

- Complete Inventory: Generates a complete list with all items registered (this button was clicked to take this print)
- Critical Items: Notice that we have the min/max/balance values. Balance is the total quantity of the item in the inventory. Min and max are values entered at the time of registering for each item, and indicate the ideal amount that should be stored. In addition, the third letter of the code for each item indicates its importance (A, B or C). By clicking on the Critical items button, we get a list of items that are below the ideal minimum and have A or B importance.
- Missing Items: generates a list with items where Balance = 0 , that is, items that are missing
- Search by Description: Search a specific item by its description
- Search by Application: Generates list of items from the chosen category.
- Search by Cod.: Searches for a specific item by its code.

3.2.2 Modifying an item

After selecting an item, we can click on:
- Inventory Receipt:

![image](https://user-images.githubusercontent.com/90487618/164837645-ddf65d3c-d8f4-46aa-9227-ac8a78df576b.png)

User had to enter the quantity to be inserted in the stock. When clicking OK the database will be updated.

- Iventory Issue:

![image](https://user-images.githubusercontent.com/90487618/164837809-da8db89e-b45c-451c-b4fc-cab7fb80e02b.png)

User had to enter the quantity to be withdrawn from the stock, when clicking OK the database will be updated.

- Edit Item

![image](https://user-images.githubusercontent.com/90487618/164837988-c2ec29fb-902a-4c80-805e-54b282d9e41f.png)

It will open a new userform with all the current information of the item, changing it and clicking on save updated the database.

- Print ID Sticker

![image](https://user-images.githubusercontent.com/90487618/164838282-8c45f6ff-ef53-4543-886d-96d9c69489d8.png)

It will generate a page formatted with the sticker tag ready for printing, as shown in the image above. The tag should be used to identify physical items. By clicking on the arrow you can return to the main menu.

3.3 New Item (Main Menu)

![image](https://user-images.githubusercontent.com/90487618/164838530-f0660109-3749-480f-83e0-a0ce5708e11c.png)

A userform for registering new items will be opened. All information must be filled in, with the exception of the code that will be generated automatically (by clicking on the button next to it). By clicking on the save button, the database will be updated with the information of the new item.

