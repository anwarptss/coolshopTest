// See https://aka.ms/new-console-template for more information
using Microsoft.VisualBasic.FileIO;
using System.Data;
using IronXL;
using coolshopTest;
using coolshopTest.DTOs;
using System.Diagnostics;

Console.WriteLine("\t\t\t------------- Welcome to the CoolShop Test! -----------------\n\n");
char rep;
do
{
    try
    {
        string customFileURL = "";
        string fileName = "File\\file.csv";

        string[] fileURL = Environment.CurrentDirectory.Split('\\');
        for (int i = 0; i < fileURL.Length - 3; i++)
        {
            customFileURL += fileURL[i] + "\\"; 
        }
        string filePath = customFileURL + fileName;//Path.Combine(Environment.CurrentDirectory, fileName);
        Common.ReadCSVData(filePath);
    }
    catch (Exception ex)
    {
        Console.WriteLine("\n Error: " + ex.Message);
    }

    Console.WriteLine("\n\n Do you want to search again? Y/N");
    rep = Convert.ToChar(Console.ReadLine());
    Console.WriteLine("\n_________________________________________________________________");

} while (rep == 'y' || rep == 'Y');

 
