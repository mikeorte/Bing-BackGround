using System;
using System.Drawing;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

public class BingBackground
{
    private const string BingApiUrl = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US";
    private const string DefaultResolution = "_1920x1080.jpg";

    private enum PicturePosition
    {
        Tile,
        Center,
        Stretch,
        Fit,
        Fill
    }

    public static void Main()
    {
        try
        {
            string backgroundUrl = GetBackgroundUrlBase() + DefaultResolution;
            Image background = DownloadBackground(backgroundUrl);
            string filePath = GetBackgroundImageFilePath();
            background.Save(filePath);
            // Set the wallpaper with the best position
            SetWallpaper(filePath, GetPosition(background));
            Console.WriteLine("Background updated successfully.");
        }
        catch (Exception ex)
        {
            // Print any errors that occur
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Method to download JSON data from Bing API
    private static dynamic DownloadJson()
    {
        try
        {
            using (WebClient webClient = new WebClient())
            {
                Console.WriteLine("Downloading JSON...");
                string jsonString = webClient.DownloadString(BingApiUrl);
                // Deserialize JSON string to dynamic object
                return JsonConvert.DeserializeObject<dynamic>(jsonString);
            }
        }
        catch (Exception ex)
        {
            // Print any errors that occur
            Console.WriteLine($"Error downloading JSON: {ex.Message}");
            return null;
        }
    }

    // Method to get the base URL for the background image
    private static string GetBackgroundUrlBase()
    {
        dynamic jsonObject = DownloadJson();
        // Construct the full URL for the image
        return "https://www.bing.com" + jsonObject.images[0].urlbase;
    }

    // Method to get the title of the background image
    private static string GetBackgroundTitle()
    {
        dynamic jsonObject = DownloadJson();
        // Extract the title from the copyright text
        string copyrightText = jsonObject.images[0].copyright;
        return copyrightText.Substring(0, copyrightText.IndexOf(" ("));
    }

    // Method to download the background image from a URL
    private static Image DownloadBackground(string url)
    {
        Console.WriteLine("Downloading background...");
        WebRequest request = WebRequest.Create(url);
        using (WebResponse response = request.GetResponse())
        using (Stream stream = response.GetResponseStream())
        {
            // Create an Image object from the stream
            return Image.FromStream(stream);
        }
    }

    // Method to get the file path to save the background image
    private static string GetBackgroundImageFilePath()
    {
        // Create a directory path based on the current year
        string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
        Directory.CreateDirectory(directory);
        // Create a file path with the current date
        return Path.Combine(directory, DateTime.Now.ToString("M-d-yyyy") + ".bmp");
    }

    // Method to set the wallpaper with a specified position
    private static void SetWallpaper(string filePath, PicturePosition position)
    {
        const int SPI_SETDESKWALLPAPER = 20;
        const int SPIF_UPDATEINIFILE = 0x01;
        const int SPIF_SENDWININICHANGE = 0x02;

        string style;
        string tileWallpaper;

        // Set the style and tileWallpaper based on the position
        switch (position)
        {
            case PicturePosition.Tile:
                style = "0";
                tileWallpaper = "1";
                break;
            case PicturePosition.Center:
                style = "0";
                tileWallpaper = "0";
                break;
            case PicturePosition.Stretch:
                style = "2";
                tileWallpaper = "0";
                break;
            case PicturePosition.Fit:
                style = "6";
                tileWallpaper = "0";
                break;
            case PicturePosition.Fill:
                style = "10";
                tileWallpaper = "0";
                break;
            default:
                style = "0";
                tileWallpaper = "0";
                break;
        }

        // Update the registry settings for the wallpaper
        using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true))
        {
            key.SetValue(@"WallpaperStyle", style);
            key.SetValue(@"TileWallpaper", tileWallpaper);
        }

        // Set the wallpaper using the SystemParametersInfo function
        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, filePath, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
    }

    // Import the SystemParametersInfo function from user32.dll
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

    // Method to determine the best position for the wallpaper
    private static PicturePosition GetPosition(Image background)
    {
        // Get screen dimensions
        int screenWidth = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
        int screenHeight = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;

        // Get image dimensions
        int imageWidth = background.Width;
        int imageHeight = background.Height;

        // Calculate aspect ratios
        float screenAspectRatio = (float)screenWidth / screenHeight;
        float imageAspectRatio = (float)imageWidth / imageHeight;

        // Determine the best position based on aspect ratios
        if (Math.Abs(screenAspectRatio - imageAspectRatio) < 0.1)
        {
            return PicturePosition.Fill; // Use Fill if aspect ratios are similar
        }
        else if (screenAspectRatio > imageAspectRatio)
        {
            return PicturePosition.Fit; // Use Fit if screen is wider than image
        }
        else
        {
            return PicturePosition.Stretch; // Use Stretch if image is wider than screen
        }
    }
}
