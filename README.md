# Template Converter

Template Converter is a tool to convert data from Excel to Word using predefined templates. The app supports transferring single cells, images, charts, and tables from Excel into Word.

## Features
- Transfer single cells, images, tables, and charts from Excel to Word.
- Customize the appearance of tables and images in Word based on Excel formatting.
- Automatically open the generated Word document (optional).

## How to Use
1. Clone the repository and build the project under the Release configuration.
2. Navigate to `TemplateConverter\src\Converter.UI\bin\Release\net6.0-windows` and open the **Converter.exe** file.
3. Select your Word template and Excel file. The final report will be generated in the same folder as the Excel file, with the same name.

### Options:
- **Open After Generated**: Check this box to automatically open the generated Word document after conversion.

## Conversion Details

### Single Cell:
- Rename the named range in Excel.
- Add a Word bookmark at the target location.
- Ensure the named range and bookmark share the same name.

### Picture:
- Place the image link in the Excel cell.
- Rename the named range in Excel.
- Add a bookmark in Word, ensuring the named range and bookmark share the same name.

### Table:
- Rename the entire table's named range in Excel.
- Add a bookmark in Word and ensure the bookmark and named range share the same name.

### Chart:
- Create a sample chart in Word.
- Ensure the chart title in Word matches the chart title in Excel.

## Requirements
- .NET 6.0 or higher
- Microsoft Word
- Microsoft Excel

## License
[Include License Information Here]
