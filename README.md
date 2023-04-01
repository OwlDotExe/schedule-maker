# ⭐ Shedule.Maker

An automation project made for saving time when it's about generating Excel activity report files.<br><br>

# Table of Contents
1. [Description](#description)
2. [Preview](#preview)
3. [Installation](#installation)<br><br>

## 📌 Project description <a name="description"></a>

<p align="justify">
Schedule.Maker is an automation project (console application) made in C# whose purpose is to generate an Excel file containing a summary of the working hours made during one specific month. This file is a sort of activity report used to formalize the work of the company's developers.

To handle the generation of the Excel file some information are needed such as :
1. The first name and the last name of the worker concerned
2. The month that has to be used for the generation
3. Potential days off taken by the worker, potential public days of the month and finally remote working days taken by the worker

For the moment, the working hours generated for each day are fixed because initialy it was for personal use and my working hours were always the same. This should be the next feature of the application to give the right to the worker to choose for each day the working hours made.

Finally some words on the libraries I used to make this project :

1. IronXL was the library that provided Excel features and Excel actions : writing cells, create workbook / worksheet, styling...
2. Spectre.Console was the library that provided console application features such as : calendar UI, progress bar, additional styling...
</p>

You can check out each library's documentation for more details : [IronXL documentation](https://ironsoftware.com/csharp/excel/features/) and [Spectre.Console documentation](https://spectreconsole.net/quick-start)<br><br>

## 👀 Project preview <a name="preview"></a>

1. User interface that summarizes all the information

![image1](https://user-images.githubusercontent.com/113787371/229347161-f9ba419a-e1f7-414c-86e9-fc48ca34d180.png)

2. Generated file by the application

![image2](https://user-images.githubusercontent.com/113787371/229347164-f863a9cf-116f-4437-8bbb-492a77ef2910.png)<br><br>

## 📃 Project installation and project use <a name="installation"></a>


1. Clone the initial repository :

```
git clone https://github.com/takoune/Schedule.Maker
```

2. You can open the SLN and launch the console application or you can go to this folder Schedule.Maker\Schedule.Maker\bin\Debug\net6.0 and launch the .exe file

3. Finally you can find your generated file in the current folder : Desktop\schedule-maker-files\name.xlsx