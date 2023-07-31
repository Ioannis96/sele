using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.ComponentModel;
using System.IO;
using OfficeOpenXml;
using Microsoft.Extensions.Configuration;
using System.Security.Cryptography.X509Certificates;
using System.Reflection.Metadata;
using System.Drawing.Text;



//ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
//string excelFilePath = @"E:\Excel\pathSpells.xlsx";
//GetLinks(excelFilePath);

SeleniumGetSpellProperties();

Console.WriteLine("telos");
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

static void SeleniumGetSpellProperties()
{
    bool hasSpellInDesc = false;

    try
    {
        var driver = new ChromeDriver();
        driver.Navigate().GoToUrl("https://www.d20pfsrd.com/magic/all-spells/e/envious-urge/");


        IWebElement descriptionElement = driver.FindElement(By.ClassName("article-text"));
        IReadOnlyList<IWebElement> namesh1 = descriptionElement.FindElements(By.TagName("h1"));
        IReadOnlyList<IWebElement> namesh4 = descriptionElement.FindElements(By.TagName("h4"));
        IReadOnlyList<IWebElement> elements = descriptionElement.FindElements(By.TagName("p"));
        List<string> spellDescriptionArray = new List<string>();

        spellDescriptionArray.Add(namesh1[0].Text);

        //if it finds an h4 element inside the 
        if (namesh4.Count > 0)
        {
            hasSpellInDesc = true;

            string name4 = namesh4[0].Text;
            Console.WriteLine("There is a spell in this spell named: " + name4);
        }

        int counter = 0;

        Console.WriteLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
        foreach (var element in elements)
        {
            spellDescriptionArray.Add(element.Text);
            Console.WriteLine( ++counter + element.Text );
        }
        Console.WriteLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");


        string? name = spellDescriptionArray[0];
        string? school = "";
        string? level = "";
        string? castingTime = "";
        string? components = "";
        string? effect = "";
        string? range = "";
        string? target = "";
        string? area = "";
        string? duration = "";
        string? savingThrow = "";
        string? description = "";

        int semicolonIndex = spellDescriptionArray[1].IndexOf(';');
        int castIndex = spellDescriptionArray[3].IndexOf("Casting Time ");

        school = spellDescriptionArray[1].Substring(0, semicolonIndex).Trim();

        level = spellDescriptionArray[1].Substring(semicolonIndex + 7).Trim();






        if (spellDescriptionArray[2].ToLower() == "casting")
        {
            string[] castArray = spellDescriptionArray[3].Split('\n');

            castingTime = castArray[0].Substring("Casting Time ".Length).Trim();

            components = castArray[1].Substring("Components ".Length).Trim();

            //why though
            if (spellDescriptionArray[4].ToLower() == "effect")
            {
                Console.WriteLine("mpenei sto effect   ");
                string[] effectArray = spellDescriptionArray[5].Split('\n');
                int len = effectArray.Length;

                for (int i = 0; i < len; i++)
                {
                    if (effectArray[i].ToLower().Contains("range"))
                    {
                        range = effectArray[i].Substring(5);
                    }
                    else if (effectArray[i].ToLower().Contains("area"))
                    {
                        area = effectArray[i].Substring(4);
                    }
                    else if (effectArray[i].ToLower().Contains("effect"))
                    {
                        effect = effectArray[i].Substring(6);
                    }
                    else if (effectArray[i].ToLower().Contains("targets"))
                    {
                        target = effectArray[i].Substring(7);
                    }
                    else if (effectArray[i].ToLower().Contains("target"))
                    {
                        target = effectArray[i].Substring(6);
                    }
                    else if (effectArray[i].ToLower().Contains("duration"))
                    {
                        duration = effectArray[i].Substring(8);
                    }
                    else if (effectArray[i].ToLower().Contains("saving throw"))
                    {

                    }
                }
            }
        }



        Console.WriteLine("Name: " + name);
        Console.WriteLine("School: " + school);
        Console.WriteLine("Level: " + level);
        Console.WriteLine("Casting Time: " + castingTime);
        Console.WriteLine("Components: " + components);
        Console.WriteLine("Range: " + range);
        Console.WriteLine("Target: " + target);
        Console.WriteLine("Duration: " + duration);
        Console.WriteLine("Saving Throw: " + savingThrow);
        Console.WriteLine("Description: " + description);


        //foreach (string element in strings)
        //{
        //        Console.WriteLine(element);
        //}

        driver.Quit();
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
        throw;
    }
}

