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
            SetWallpaper(filePath, GetPosition(background));
            Console.WriteLine("Background updated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static dynamic DownloadJson()
    {
        try
        {
            using (WebClient webClient = new WebClient())
            {
                Console.WriteLine("Downloading JSON...");
                string jsonString = webClient.DownloadString(BingApiUrl);
                return JsonConvert.DeserializeObject<dynamic>(jsonString);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error downloading JSON: {ex.Message}");
            return null;
        }
    }

    private static string GetBackgroundUrlBase()
    {
        dynamic jsonObject = DownloadJson();
        return "https://www.bing.com" + jsonObject.images[0].urlbase;
    }

    private static string GetBackgroundTitle()
    {
        dynamic jsonObject = DownloadJson();
        string copyrightText = jsonObject.images[0].copyright;
        return copyrightText.Substring(0, copyrightText.IndexOf(" ("));
    }

    private static Image DownloadBackground(string url)
    {
        Console.WriteLine("Downloading background...");
        WebRequest request = WebRequest.Create(url);
        using (WebResponse response = request.GetResponse())
        using (Stream stream = response.GetResponseStream())
        {
            return Image.FromStream(stream);
        }
    }

    private static string GetBackgroundImageFilePath()
    {
        string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, DateTime.Now.ToString("M-d-yyyy") + ".bmp");
    }

    private static void SetWallpaper(string filePath, PicturePosition position)
    {
        const int SPI_SETDESKWALLPAPER = 20;
        const int SPIF_UPDATEINIFILE = 0x01;
        const int SPIF_SENDWININICHANGE = 0x02;

        string style;
        string tileWallpaper;

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

        using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true))
        {
            key.SetValue(@"WallpaperStyle", style);
            key.SetValue(@"TileWallpaper", tileWallpaper);
        }

        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, filePath, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

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

