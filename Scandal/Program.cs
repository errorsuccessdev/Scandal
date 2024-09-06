﻿/*
 * TODO:
 *  - Better input handling
 *  - Nicer UI :)
 *  - Better solution for selecting PDF creation method
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using WIA;
using Image = System.Drawing.Image;

namespace Scandal
{
    internal class Program
    {
        const int WIA_S_NO_DEVICE_AVAILABLE = -2145320939;

        static void Main()
        { 
            bool noScannersFound = true;
            DeviceManager deviceManager = new DeviceManager();
            DeviceInfos devices = deviceManager.DeviceInfos;

            foreach (DeviceInfo device in devices)
            {
                if (device.Type ==
                    WiaDeviceType.ScannerDeviceType)
                {
                    noScannersFound = false;
                    Console.Write(
                        "Press space to scan, p to print, or q to quit: "
                    );
                    List<Image> images = new List<Image>();
                    while (true)
                    {
                        ConsoleKeyInfo key = Console.ReadKey(true);
                        if (key.Key == ConsoleKey.Spacebar)
                        {
                            images.Add(scan(device));
                        }
                        else if (key.KeyChar == 'p' &&
                            images.Count > 0)
                        {
                            print(images);
                        }
                        else if (key.KeyChar == 'q')
                        {
                            break;
                        }
                    }
                }
            }

            if (noScannersFound)
            {
                Console.WriteLine("No scanners were found.");
            }
        }

        static Image scan(DeviceInfo info)
        {
            CommonDialog dialog = new CommonDialog();
            Device device = info.Connect();
            Image image = null;
            try
            {
                ImageFile imageFile;

                // This should probably not be hard coded
                imageFile = dialog.ShowTransfer(device.Items[1]);

                if (imageFile != null)
                {
                    image = toImage(imageFile);
                }
            }
            catch (COMException ex)
            when (ex.ErrorCode == WIA_S_NO_DEVICE_AVAILABLE)
            {
                // handle no devices available
                Console.WriteLine("No scanner devices were found.");
            }
            return image;
        }

        static Image toImage(ImageFile imageFile)
        {
            MemoryStream stream = new MemoryStream();
            Image image = null;
            try
            {
                byte[] data;
                data = (byte[]) imageFile.FileData.get_BinaryData();
                stream.Write(
                    data, 
                    0, 
                    data.Length
                );
                image = Image.FromStream(stream);
            }
            catch
            {
                Console.WriteLine(
                    "There was a problem getting the image."
                );
            }
            finally
            {
                // What is going on here? What does this do? Why do we need it?
                stream.Dispose();
            }
            return image;
        }

        static Bitmap toBitmap(Image image)
        {
            Bitmap result = null;
            MemoryStream stream = new MemoryStream();
            Image scannedImage = null;
            try
            {
                result = new Bitmap(
                        image.Width,
                        image.Height,
                        PixelFormat.Format32bppArgb
                    );
                result.Save(
                    "test.png",
                    ImageFormat.Png
                );
                List<Image> images = 
                    new List<Image> { scannedImage, scannedImage };
                print(images);
            }
            catch
            {
                Console.WriteLine(
                    "There was a problem converting the bitmap."
                );
            }
            finally
            {
                // What is going on here? What does this do? Why do we need it?
                stream.Dispose();
            }
            return result;
        }

        static void print(List<Image> images)
        {
            try
            {
                PrintDocument printDocument = new PrintDocument();
                printDocument.PrinterSettings.PrinterName =
                    "Microsoft Print to PDF";
                printDocument.PrintPage +=
                    delegate (object sender, PrintPageEventArgs eventArgs)
                    {
                        pd_PrintPage(
                            sender,
                            eventArgs, 
                            images
                        );
                    };
                if (printDocument.PrinterSettings.IsValid)
                {
                    printDocument.Print();
                }
                else
                {
                    Console.WriteLine("Invalid printer settings!");
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        // The PrintPage event is raised for each page to be printed
        static private void pd_PrintPage(
            object sender,
            PrintPageEventArgs eventArgs,
            List<Image> images
        )
        {
            eventArgs.Graphics.DrawImage(
                images[0], 
                eventArgs.PageBounds
            );
            images.RemoveAt(0);
            eventArgs.HasMorePages = (images.Count > 0);
        }
    }
}