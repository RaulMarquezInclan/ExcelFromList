# ExcelFromList
Straightforward and easy way to create stylized excel workbooks from lists. Add an image, title, subtitles and overal cell styles/formats. This uses the EPPlus engine, you can check them out at: <br />
https://github.com/EPPlusSoftware/EPPlus

In the below examples <b>outputFileName</b> is a string identifying a full path file name and <b>shelfLifeData</b> is a list of <b>ShelfLife</b> type objects.

You can run these same examples in the <b>Testing</b> project.

## With default styles (no style object provided)
```
  var wb = new ExcelWorkBook();
  wb.AddSheet("Shelf Life", shelfLifeData);
  wb.SaveAs(outputFileName);
```
![Default styles](https://i.imgur.com/MwOVdeQ.png)

A style object has been provided for the rest of the examples
## With title and subtitles
```
  var wb = new ExcelWorkBook();
  var style = new ExcelStyleConfig
  {
		Title = "Product Shelf Life List",
		Subtitles = new string[]
		{
			"As of 2/1/06",
			"Compiled by the Food Bank",
			"From National Manufactures"
		}
  };

  wb.AddSheet("Shelf Life", shelfLifeData, style);
  wb.SaveAs(outputFileName);
```
![Title and subtitles](https://i.imgur.com/sBWGHrM.png)

## With title, subtitles and image from Base64
```
  var wb = new ExcelWorkBook();
  var style = new ExcelStyleConfig
  {
		Title = "Product Shelf Life List",
		Subtitles = new string[]
		{
			"As of 2/1/06",
			"Compiled by the Food Bank",
			"From National Manufactures"
		},
		TitleImage = new Picture()
		{
			FromBase64 = "iVBORw..." // string trucated for brevity of example
		}
  };

  wb.AddSheet("Shelf Life", shelfLifeData, style);
  wb.SaveAs(outputFileName);
```
![Title, subtitles and image](https://i.imgur.com/vEJp6Yx.png)

## With title, subtitles, image from file (sheetOneStyle) and url (sheetTwoStyle), two sheets and cell stylings
```
  var wb = new ExcelWorkBook();
  var sheetOneStyle = new ExcelStyleConfig
  {
      Title = "Product Shelf Life List",
      Subtitles = new string[]
      {
				"As of 2/1/06",
				"Compiled by the Food Bank",
				"From National Manufactures"
      },
      TitleImage = new Picture()
      {
				FromFile = @"x:\titleImage.jpg"
      }
  };
  var sheetTwoStyle = new ExcelStyleConfig
  {
      Title = "Food Nutrient Information",
      Subtitles = new string[]
      {
				"List of EDNP products",
				"Audited by category"
      },
      TitleImage = new Picture()
      {
				FromUrl = @"http://www.images.com/titleImage.jpg"
      },
      ShowGridLines = false,
      BorderAround = true,
      Border = true,
      BorderColor = Color.CadetBlue,
      HeaderBackgroundColor = Color.Yellow,
      HeaderFontColor = Color.Black
  };

  wb.AddSheet("Shelf Life", shelfLifeData, sheetOneStyle);
  wb.AddSheet("Food Nutrients", foodInfoData, sheetTwoStyle);
  wb.SaveAs(outputFileName);
```
![Title, subtitles, image and two sheets](https://i.imgur.com/LpDg2pb.png)

## With title, subtitles, image from Base64, skipping two columns and cell stylings
```
var wb = new ExcelWorkBook();
var style = new ExcelStyleConfig
{
	Title = "Food Nutrient Information",
	Subtitles = new string[]
	{
		"List of EDNP products",
		"Audited by category"
	},
	TitleImage = new Picture()
	{
		FromBase64 = "iVBORw..." // string trucated for brevity of example
	},
	ShowGridLines = false,
	BorderAround = true,
	Border = true,
	BorderColor = Color.CadetBlue,
	HeaderBackgroundColor = Color.Yellow,
	HeaderFontColor = Color.Black,
	ExcludedColumnIndexes = new int[]
	{
		2, 4
	}
};

wb.AddSheet("Food Nutrients", shelfLifeData, style);
wb.SaveAs(outputFileName);
```
![Title, subtitles, image, skipping three rows and cell stylings](https://i.imgur.com/BFid7jk.png)

# Documentation
## Available in the ExcelStyleConfig class
### Sheet configs
<b>ShowHeaders:</b> Enable to show headers (taken from the property name), defaults to <b>true</b><br />
<b>ShowGridLines:</b> Enable to show grid lines, defaults to <b>true</b><br />
<b>AutoFitColumns:</b> Enable to match the width of the column to the data length, defaults to <b>true</b><br />
<b>FreezePanes:</b> Enable to freeze the first row, defaults to <b>true</b><br />
<b>PaddingColumns:</b> Gets or sets the number of columns to insert before column A, defaults to <b>0</b><br />
<b>PaddingRows:</b> Gets or sets the number of rows to insert before row 1, defaults to <b>0</b><br />
<b>ExcludedColumnIndexes:</b> Gets or sets which columns to exclude by index, range must be between 1 and the total number of columns, defaults to <b>new int[0]</b><br />

### Title configs
<b>Title:</b> Gets or sets the title of the sheet, defaults to <b>null</b><br />
<b>Subtitles:</b> Gets or sets the subtitles of the sheet, defaults to <b>new string[0]</b><br />
<b>TitleImage:</b> Gets or sets an image to be placed on the sheet, defaults to <b>new Picture()</b><br />
<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FromBase64</b>: Gets or sets image from Base64, defaults to <b>null</b><br />
<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FromFile</b>: Gets or sets image from file, defaults to <b>null</b><br />
<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FromUrl</b>: Gets or sets image from url, defaults to <b>null</b><br />
<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;HasValue</b>: Indicates if at least one image source has value, defaults to <b>false</b><br />
<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;IsValid</b>: Indicates if value is valid, no source or more than one source (false), only one source (true), defaults to <b>false</b><br />

### Data type formatting
<b>DateFormat:</b> Gets or sets custom Excel format string, defaults to <b>m/d/yyyy</b><br />
<b>DecimalFormat:</b> Gets or sets custom Excel format string, defaults to <b>#,##0.00_);[Red]-#,##0.00</b><br />
<b>DoubleFormat:</b> Gets or sets custom Excel format string, defaults to <b>#,##0.00_);[Red]-#,##0.00</b><br />
<b>IntFormat:</b> Gets or sets custom Excel format string, defaults to <b>#,##0_);[Red]-#,##0</b><br />
<i>* Uses Excel type data formatting</i><br />

### Data cell configs
<b>FontColor:</b> Gets or sets data cell font color, defaults to <b>null</b><br />
<b>BackgroundColor:</b> Gets or sets data cell background color, defaults to <b>null</b><br />
<b>Border:</b> Enable to draw a border around each data cell, defaults to <b>false</b><br />
<b>BorderAround:</b> Enable to draw a border around the data range, defaults to <b></b>false<br />
<b>BorderColor:</b> Gets or sets the border color around each data cell, defaults to <b>Color.Black</b><br />
<b>BorderAroundColor:</b> Gets or sets the border color around the data range, defaults to <b>Color.Black</b><br />
<b>BorderStyle:</b> Gets or sets the border style around each data cell, defaults to <b>ExcelBorderStyle.Thin</b><br />
<b>BorderAroundStyle:</b> Gets or sets the border style around the data range, defaults to <b>ExcelBorderStyle.Thin</b><br />

### Header cell configs
<b>HeaderFontColor:</b> Gets or sets the header font color, defaults to <b>Color.LightGray</b><br />
<b>HeaderBackgroundColor:</b> Gets or sets the header background color, defaults to <b>Color.DarkSlateGray</b><br />
<b>HeaderBorder:</b> Enable to draw a border around each header cell, defaults to <b>false</b><br />
<b>HeaderBorderAround:</b> Enable to draw a border around the header range, defaults to <b>false</b><br />
<b>HeaderBorderColor:</b> Gets or sets the border color around each header cell, defaults to <b>Color.Black</b><br />
<b>HeaderBorderAroundColor:</b> Gets or sets the border color around the header range, defaults to <b>Color.Black</b><br />
<b>HeaderBorderStyle:</b> Gets or sets the border style around each header cell, defaults to <b>ExcelBorderStyle.Thin</b><br />
<b>HeaderBorderAroundStyle:</b> Gets or sets the border style around the header range, defaults to <b>ExcelBorderStyle.Thin</b><br />

## Available in the ExcelWorkBook class
### Methods
<b>GetBytesArray():</b> Returns the ExcelWorkBook bytes array<br />
<b>AddSheet(string sheetName):</b> Adds a sheet to the worksheet, will apply style config if provided<br />
<b>RemoveSheet(string sheetName):</b> Removes a sheet from the worksheet<br />
<b>ClearWorkSheet():</b> Removes all sheets from worksheet<br />
<b>SheetExists(string sheetName):</b> Checks if a specific sheet exists<br />
<b>SaveAs():</b> Saves the workbook to an Excel file<br />
<b>Open():</b> Opens saved Excel file with OS default program<br />

For new features you can contact me at raulmarquezi@gmail.com
