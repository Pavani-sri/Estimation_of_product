# Estimation_of_product
This project is built when learing about modern Excel automation
In the Main.xaml  

Create a new sequence.
Provide an annotation for the Main sequence. 
Add the Excel Process Scope activity to load the Excel project settings and apply them to the associated Excel file.
Add the Use Excel activity to select the "VegetablesProduction.xlsx" Excel file to use in the automation. 
Configure it with the following properties:
Excel File : "VegetablesProduction.xlsx
Create if not exist : checked  
Inside the DO container of the Use Excel activity, add the following activities: 
Add the Format Cells activity to format the "Estimated Value" column to currency and add the currency sign.
Add the Create Pivot Table activity and configure it to have "Vendor Name" and "Production Month" as rows, "Storage Zone" as columns, and "Estimated Value" as a value using the SUM function.
Add the Insert Chart activity to create a Stacked Column Chart based on the Pivot Table information in the PivotTable Sheet.
Add the Update Excel Chart activity to update the chart name to "Estimated Sales".
