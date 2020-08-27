﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
//using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Word;

namespace PowerSecurePDFPOC
{
    class Program
    {
        static void Main(string[] args)
        {
            AppContext.SetSwitch("Switch.System.Xml.AllowDefaultResolver", true);
            var normalizedDocType = "";
            var pause = false;
            if (args != null && args.Count() > 0)
            {
                normalizedDocType = args[0].Trim().Replace(" ", "");
            }
            if (normalizedDocType == "doctype=template")
            {
                var inputTemplateFile = $"./WorkflowDocuments/input/PowerSecure-OwnedAsset-Template-Instrumented.xlsx";
                var outputFlattendOPCTemplateFile = $"./WorkflowDocuments/output/FlattenedOPC/PowerSecure-OwnedAsset-Template-Instrumented-Flattened.xml";
                echo($"inputTemplate File Name: {inputTemplateFile}");
                var templateSpreadsheet = SpreadsheetDocument.Open(inputTemplateFile, false);
                var flattenedTemplateXDoc = templateSpreadsheet.ToFlatOpcDocument();
                flattenedTemplateXDoc.Save(outputFlattendOPCTemplateFile);
                echo($"input templateFile name: {outputFlattendOPCTemplateFile}", pause);
            }
            else
            {
                var inputJSONFile = $"./WorkflowDocuments/input/Summary-Response-MultiLineItem.json";
                var xsltFile = @"./assets/templates/OwnedAssetReport_MultiLine/PowerSecure-OwnedAsset-Templatexml.xslt";
                var outputJSONAsXmlFile = $"./WorkflowDocuments/output/Intermediate/Summary-Response-20200818-{DateTime.Now.Ticks.ToString()}.xml";
                var outputFlattenedOPCFile = $"./WorkflowDocuments/output/FlattenedOPC/OwnedAsset-FlattenedOPC-{DateTime.Now.Ticks.ToString()}.xml";
                var outputExcelFile = $"./WorkflowDocuments/output/StandardOPC/OwnedAsset-{DateTime.Now.Ticks.ToString()}.xlsx";

                var inputXMLDoc = new XmlDocument();

                echo($"input JSON file name: {inputJSONFile}", pause);
                echo($"input xslt file name: {xsltFile}", pause);
                echo($"output intermediate XML Data file name: {outputFlattenedOPCFile}", pause);
                echo($"output flattened OPC file name: {outputFlattenedOPCFile}", pause);
                echo($"output Excel file name: {outputExcelFile}", pause);

                using (StreamReader jsonReader = new StreamReader(inputJSONFile))
                {
                    string inputJSONString = jsonReader.ReadToEnd();
                    inputXMLDoc = (XmlDocument)JsonConvert.DeserializeXmlNode(inputJSONString);
                };

                echo($"Create a writer for the output of the Xsl Transformation...", pause);
                //Create a writer for the output of the Xsl Transformation.
                Utf8StringWriter stringWriter = new Utf8StringWriter(Encoding.UTF8);
                var xmlSettings = new XmlWriterSettings();
                xmlSettings.Encoding = Encoding.UTF8;
                XmlWriter xmlWriter = XmlWriter.Create(stringWriter, xmlSettings);

                echo($"Create the Xsl Transformation object...", pause);
                //Create the Xsl Transformation object.
                XslCompiledTransform transform = new XslCompiledTransform();
                XsltSettings settings = new XsltSettings(true, true);
                var resolver = new XmlUrlResolver();
                transform.Load(xsltFile, settings, resolver);

                echo($"Transform the xml data into Open XML 2.0 Spreadsheet format...", pause);
                inputXMLDoc.Save(outputJSONAsXmlFile);
                var inputDocReader = new XmlTextReader(outputJSONAsXmlFile);
                //Transform the xml data into Open XML 2.0 SpreadSheet format FROM XML FILE.
                //transform.Transform(xmlDataFile, xmlWriter);
                //Transform the xml data into Open XML 2.0 SpreadSheet format FROM XML parsed from input JSON file.
                transform.Transform(inputDocReader, xmlWriter);
                inputDocReader.Close();

                echo($"Create an Xml Document of the new content...", pause);
                //Create an Xml Document of the new content.
                XmlDocument newSpreadsheetContent = new XmlDocument();
                newSpreadsheetContent.LoadXml(stringWriter.ToString());

                echo($"Output Xml document to file...", pause);
                // Create a new file for the transformed content in flattened OPC format
                newSpreadsheetContent.Save(outputFlattenedOPCFile);

                echo($"Convert the Flattened OPC document to Standard OPC and output document to disk...", pause);
                // Convert the flattened document output to Standard OPC for output as normal XLSX file
                //var xDocument = System.Xml.Linq.XDocument.Load(outputFlattenedOPCFile);
                //var spreadsheet = SpreadsheetDocument.FromFlatOpcDocument(xDocument);
                var spreadsheet = SpreadsheetDocument.FromFlatOpcString(stringWriter.ToString());
                var package = spreadsheet.SaveAs(outputExcelFile);
                package.Close();
                xmlWriter.Close();
                stringWriter.Close();
                spreadsheet.Close();

                // Now lets validate the results
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    int count = 0;
                    foreach (ValidationErrorInfo error in validator.Validate(SpreadsheetDocument.Open(outputExcelFile, true)))
                    { 
                        count++;
                        Console.WriteLine("Error " + count);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }
                    //Console.ReadKey();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message); 
                }
            }
        }

        public static void echo(string message, bool pause = false)
        {
            Console.WriteLine(message);
            if (pause)
            {
                Console.WriteLine("Press any key...");
                Console.ReadKey();
            }
        }
    }
}
