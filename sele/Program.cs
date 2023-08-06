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
using OpenQA.Selenium.DevTools;
using System.Text;
using System.Xml.Linq;



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


    string? name = "";
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
    string? spellResistance = "";
    string? description = "";

    try
    {
        var driver = new ChromeDriver();
        driver.Navigate().GoToUrl("https://www.d20pfsrd.com/magic/all-spells/a/acid-pit/");


        IWebElement descriptionElement = driver.FindElement(By.ClassName("article-text"));
        IReadOnlyList<IWebElement> namesh1 = descriptionElement.FindElements(By.TagName("h1"));
        IReadOnlyList<IWebElement> namesh4 = descriptionElement.FindElements(By.TagName("h4"));
        IReadOnlyList<IWebElement> elements = descriptionElement.FindElements(By.TagName("p"));
        List<string> spellDescriptionArray = new List<string>();

        spellDescriptionArray.Add(namesh1[0].Text);

        //if it finds an h4 element inside the 
        //if (namesh4.Count > 0)
        //{
        //    hasSpellInDesc = true;

        //    string name4 = namesh4[0].Text;
        //    Console.WriteLine("There is a spell in this spell named: " + name4);
        //}

        
        foreach (var element in elements)
        {
            spellDescriptionArray.Add(element.Text);
        }

        name = spellDescriptionArray[0];

        int semicolonIndex = spellDescriptionArray[1].IndexOf(';');
        int castIndex = spellDescriptionArray[3].IndexOf("Casting Time ");
        school = spellDescriptionArray[1].Substring(0, semicolonIndex).Trim();
        level = spellDescriptionArray[1].Substring(semicolonIndex + 7).Trim();
        int propLen = spellDescriptionArray.Count();

        for(int i=0; i<propLen; i++)
        {
            if (spellDescriptionArray[i].ToLower() == "casting")
            {
                string[] castArray = spellDescriptionArray[i+1].Split('\n');
                castingTime = castArray[0].Substring(14).Trim();
                components = castArray[1].Substring(11).Trim();
            }
            else if (spellDescriptionArray[i].ToLower() == "effect")
            {
                Console.WriteLine("mpenei sto effect   ");
                string[] effectArray = spellDescriptionArray[i+1].Split('\n');
                int len = effectArray.Length;

                for (int j = 0; j < len; j++)
                {
                    if (effectArray[j].ToLower().Contains("range"))
                    {
                        range = effectArray[j].Substring(5).Trim();
                    }
                    else if (effectArray[j].ToLower().Contains("area"))
                    {
                        area = effectArray[j].Substring(4).Trim();
                    }
                    else if (effectArray[j].ToLower().Contains("effect"))
                    {
                        effect = effectArray[j].Substring(6).Trim();
                    }
                    else if (effectArray[j].ToLower().Contains("targets"))
                    {
                        target = effectArray[j].Substring(7).Trim();
                    }
                    else if (effectArray[j].ToLower().Contains("target"))
                    {
                        target = effectArray[j].Substring(6).Trim();
                    }
                    else if (effectArray[j].ToLower().Contains("duration"))
                    {
                        duration = effectArray[j].Substring(8).Trim();
                    }
                    else if (effectArray[j].ToLower().Contains("saving throw"))
                    {
                        int index = effectArray[j].IndexOf("Spell Resistance");
                        if (index != -1)
                        {
                            string saveString = effectArray[j].Substring(0, index).Trim();
                            string SRString = effectArray[j].Substring(index).Trim();

                            savingThrow = saveString;
                            spellResistance = SRString;
                        }
                    }
                }
            }
            else if (spellDescriptionArray[i].ToLower() == "description")
                {
                    description = spellDescriptionArray[i+1];
                }
        }




        Console.WriteLine(" KSEKINIMA ");
        Console.WriteLine("Name: " + name);
        Console.WriteLine("School: " + school);
        Console.WriteLine("Level: " + level);
        Console.WriteLine("Casting Time: " + castingTime);
        Console.WriteLine("Components: " + components);
        Console.WriteLine("Range: " + range);
        Console.WriteLine("Target: " + target);
        Console.WriteLine("Duration: " + duration);
        Console.WriteLine("Saving Throw: " + savingThrow);
        Console.WriteLine("Spell Resistance: " + spellResistance);
        Console.WriteLine("Description: " + description);


        driver.Quit();
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
        throw;
    }

    
}

