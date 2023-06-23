# Automatize with VBA a recurring analysis
The aim of this project is to automatize with VBA a typical weekly or montly analysis where the structure of data is always the same.
Suppose for example that every Monday, on a weekly basis, you have to download data from a software internal to your company and perform the same analysis process (e.g cleaning, preparation and analysis).
With the Vba_project file in the repository, it is possible to automatize the process simply by inserting the file path and name. The user is allowed to:

* **Load data and perform the task**. The latter consists of: data cleaning and preparation (correct the data type and create relevant columns such as % of margin for each product)
* **Create a chart with the historical performance on a selected KPI**: the default kpi is 'Units Sold' but the user is allowed to select the preferred one using the listbox. Each time the kpi is changed, the programme ask the user if he want to save a copy of the file in the same folder
* **Create a raw ppt file with the graph of the selected historical data**: the ppt file is a very basic one, as a starting point for a more designed one.
* **Send an email to the person you usually need to report to with the excel and ppt file**: the user is allowed to use the text box in the file in order to write the text of the mail he wants to send

All this actions can be performed simply by using the buttons provided. An user should firstly insert the file path, secondly run the analysis and ultimately decide to create the powerpoint and send the email.
However, the error handling procedures can manage an user that does not follow the instructions, asking him if he wants to perform the previous actions before the selected one.
