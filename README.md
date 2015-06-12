# CcExcelWriter - Clean Code Excel Writer
Easiest library written in C# to write excel files (2007 and later). With this library you can write excel files with extremely low amount of code and configure fancy excel styles with no knowledge of how to do it in C#.

## Introduction
DocumentFormat.OpenXML is one of the best libraries to work with excel files. But if you already use it or is learning how to do it, you will notice that this library are very complex and requires you to have a extensive knowledge about excel files structure.

CcExcelWriter is a extension of the library DocumentFormat.OpenXML, that fix all common problems of write excel files. So you have all the power of OpenXML but don't need complex code to do it.

**Get and Create sheet, lines and cells** become more easy, the only thing you do is specified sheet name, and cells id (column and line like A1, A2, B1, B2, ...)

**Write data of different types**, is simple now. This library resolve all problems to you, you can write all primitive types and other like String, DateTime, TimeSpan, without thinking of how they will be placed in excel file.

**Creating styles in cells**, it's a big complex thing to do in OpenXML, but with this library you don't need to think in nothing of that. You can simple create an excel file and copy the cell styles from it.

So check it out the code below to get a glimpse of what CcExcelWrite can do.

## Finding the binaries

It is recommended to use Nuget
```
Install-Package CcExcelWriter
```

## How to Use - Sample 01

Let's do a simple example of how to use CcExcelWriter, this exemple it isn't functional (don't have a real application) but represents well the basic resources of this library.

For start let's make a excel file like that:

<img src='http://s22.postimg.org/z8b5vowgx/Cc_Excel_Write_Sample01_01.png' />

Nothing special, the cells B2, C2 and D2 has font-family, font-size, background-color, font-color, text-aling, and border defined.

Now let's do some **C# code**:

```c#
// Get a stream of the existing excel file
// You can use any kind of stream, don't need to be FileStream.
using (var fs = new FileStream("fileName", FileMode.Open))
{
    // Open the excel to work.
    var excel = new Excel(fs);

    // Get the sheet (Plan1 is the name of the excel existing sheet,
    // to more information see comment below.
    var sheet = excel.GetSheet("Plan1");

    // Copy the styles of the existing cells to use later.
    var styleR = sheet.GetCellStyle("B", 2);
    var styleG = sheet.GetCellStyle("C", 2);
    var styleB = sheet.GetCellStyle("D", 2);

    // Clean all that exists in Plan1 to make something new.
    sheet.ClearAllSheetData();

    // Write information in some cells using the styles
    // that was copy before.
    sheet.SetCell("B", 2, styleR, "Some");
    sheet.SetCell("D", 2, styleG, "Simple");
    sheet.SetCell("F", 2, styleB, "Text");
    sheet.SetCell("B", 4, styleR, 10);
    sheet.SetCell("D", 4, styleG, 12.3);
    sheet.SetCell("F", 4, styleB, '-');

    // Save all updates.
    excel.Save();
}
```

*The `Plan1` name used in GetSheet is the same of the excel, see the image below:  
<img src='http://s24.postimg.org/a7ar6wznp/Cc_Excel_Write_Sample01_02.png' />

The image below is the result of the code:  
<img src='http://s22.postimg.org/i0eqkql41/Cc_Excel_Write_Sample01_03.png' />

## Sample 02

Let's do something more interesting and functional, an employee data sheet. So first we make a excel like: 

<img src="http://s28.postimg.org/ohmdxkujh/Cc_Excel_Write_Sample02_01.png" />

Then we need get data from some database. I won't specify how this can be done it's not the point here. Instead know there's a previous code that load data in the `employees` variable. Assume that variable employees implements `IEnumerable<Employee>` Where:

```C#
public class Employee
{
    public long ID { get; set; }
    public string Name { get; set; }
    public string Email { get; set; }
    public DateTime Birthday { get; set; }
    public decimal Salary { get; set; }
    public DateTime StartDate { get; set; }
}
```

So the code will be like:

```C#
using (var fs = new FileStream(fileOut, FileMode.Open))
{
    var excel = new Excel(fs);
    var sheet = excel.GetSheet("Plan1");

    // Copy the styles of the existing cells to use later.
    var styleInt = sheet.GetCellStyle("A", 2);
    var styleFloat = sheet.GetCellStyle("E", 2);
    var styleString = sheet.GetCellStyle("B", 2);
    var styleDate = sheet.GetCellStyle("D", 2);

    // Clean all data of line 2 and below. Leave the header intact.
    sheet.ClearAllSheetDataBelow(2);

    uint line = 2;

    foreach (var employee in employees)
    {
        sheet.SetCell("A", line, styleInt, employee.ID);
        sheet.SetCell("B", line, styleString, employee.Name);
        sheet.SetCell("C", line, styleString, employee.Email);
        sheet.SetCell("D", line, styleDate, employee.Birthday);
        sheet.SetCell("E", line, styleFloat, employee.Salary);
        sheet.SetCell("F", line, styleDate, employee.StartDate);

        line++;
    }

    // Save all updates.
    excel.Save();
}
```

*Notice that you don't need to copy every column style, do column D and F the style of date is the same so you can use it to both.

The result of this code (Depending on the data you get from your database) will be something like that:

<img src="http://s28.postimg.org/d93lw1rbx/Cc_Excel_Write_Sample02_02.png" />
