using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.ComponentModel;
using System.IO;
using OfficeOpenXml;
using Microsoft.Extensions.Configuration;


//ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
//string excelFilePath = @"E:\Excel\pathSpells.xlsx";
//GetLinks(excelFilePath);

Sele();

Console.WriteLine(  "telos");
Console.ReadLine();



static void GetLinks(string excelFilePath)
{
    try
    {
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["spellsWithLinks"];

            int rowCount = worksheet.Dimension.Rows;
            int hyperlinkColumnIndex = 1; //the index of the column.

            for (int row = 1; row <= rowCount; row++)
            {
                var cell = worksheet.Cells[row, hyperlinkColumnIndex];

                // Check if the cell contains a hyperlink
                if (cell.Hyperlink != null)
                {
                    string hyperlinkUrl = cell.Hyperlink.ToString();

                    Console.WriteLine($"Hyperlink in Row {row}, Column {hyperlinkColumnIndex}: {hyperlinkUrl}");
                }
            }
        }


    }
    catch (Exception ex)
    {
        Console.WriteLine("Error accessing the Excel file: " + ex.Message);
    }
}

static void Sele()
{
    try
    {
        var driver = new ChromeDriver();
        driver.Navigate().GoToUrl("https://www.d20pfsrd.com/magic/all-spells/a/align-weapon/");

        //driver.Navigate().GoToUrl(
        //IReadOnlyList<IWebElement> spells = driver.FindElements(By.TagName("li"));
        //Console.WriteLine(spells.Count());
        //foreach (var spell in spells)
        //{
        //    Console.WriteLine(spell.Text);
        //}

        IReadOnlyList<IWebElement> namesh1 = driver.FindElements(By.TagName("h1"));
        IReadOnlyList<IWebElement> namesh4 = driver.FindElements(By.TagName("h4"));

        foreach (var h4 in namesh4)
        {
            Console.WriteLine(h4.Text);
        }
        Console.WriteLine(namesh1.Count());
        Console.WriteLine(namesh4.Count());


        IWebElement descriptionElement = driver.FindElement(By.ClassName("article-text"));

        IReadOnlyList<IWebElement> elements = descriptionElement.FindElements(By.TagName("p"));

        foreach (var element in elements)
        {
            Console.WriteLine(element.Text);
        }


        //IReadOnlyList<IWebElement> elements = driver.FindElements(By.TagName("p"));
        //foreach (IWebElement e in elements)
        //{
        //    System.Console.WriteLine(e.Text);
        //}


        driver.Quit();
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
        throw;
    }
}