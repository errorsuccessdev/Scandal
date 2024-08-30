/*
 * TODO:
 *  - Better input handling
 *  - Nicer UI :)
 *  - Hardcode Microsoft Print to PDF as printer instead of relying
 *      on default printer.
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
            DeviceManager deviceManager;
            DeviceInfos devices;
            bool noScannersFound = true;

            deviceManager = new DeviceManager();
            devices = deviceManager.DeviceInfos;

            foreach (DeviceInfo device in devices)
            {
                if (device.Type ==
                    WiaDeviceType.ScannerDeviceType)
                {
                    Console.Write("Press enter to scan, or p to print: ");
                    noScannersFound = false;
                    List<Image> images = new List<Image>();
                    while (true)
                    {
                        ConsoleKeyInfo key = Console.ReadKey(true);
                        if (key.Key == ConsoleKey.Enter)
                        {
                            images.Add(scan(device));
                        }
                        else if (key.KeyChar == 'p')
                        {
                            break;
                        }
                    }

                    print(images);
                }
            }

            if (noScannersFound)
            {
                Console.WriteLine("No scanners were found.");
            }

            Console.Read();
        }

        static Image scan(DeviceInfo info)
        {
            CommonDialog dialog = new CommonDialog();
            Device device = info.Connect();
            Image image = null;
            try
            {
                ImageFile imageFile;

                imageFile = dialog.ShowTransfer(device.Items[1]);

                if (imageFile != null)
                {
                    // do something with the image
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
                data = (byte[])imageFile.FileData.get_BinaryData();
                stream.Write(data, 0, data.Length);
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

        // The Click event is raised when the user clicks the Print button.
        static void print(List<Image> images)
        {
            try
            {
                PrintDocument pd = new PrintDocument();
                pd.PrintPage +=
                    delegate (object sender, PrintPageEventArgs ev)
                    {
                        pd_PrintPage(sender, ev, images);
                    };
                pd.Print();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        // The PrintPage event is raised for each page to be printed.
        static private void pd_PrintPage(
            object sender,
            PrintPageEventArgs ev,
            List<Image> images
        )
        {
            ev.Graphics.DrawImage(images[0], ev.PageBounds);
            images.RemoveAt(0);
            ev.HasMorePages = (images.Count > 0);
        }
    }
}