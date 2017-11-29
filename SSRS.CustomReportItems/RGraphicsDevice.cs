using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using Microsoft.Win32;

/// <summary>
///     Captures the graphics device output of an arbitrary R code into an <see cref="System.Drawing.Image" /> object.
/// </summary>
public class RGraphicsDevice
{
    /// <summary>
    ///     Constructor.
    /// </summary>
    public RGraphicsDevice(double deviceWidthInMillimeters, double deviceHeightInMillimeters,
        float deviceResolutionInDpi, bool renderTextOutputInstead, Font textOutputFont, Brush textOutputBrush)
    {
        // adopt device width, height and resolution
        DeviceWidthInMillimeters = deviceWidthInMillimeters;
        DeviceHeightInMillimeters = deviceHeightInMillimeters;
        DeviceResolutionInDpi = deviceResolutionInDpi;
        RenderTextOutputInstead = renderTextOutputInstead;
        TextOutputFont = textOutputFont;
        TextOutputBrush = textOutputBrush;
    }

    /// <summary>
    ///     Gets or sets the device's width in millimeters.
    /// </summary>
    public double DeviceWidthInMillimeters { get; set; }

    /// <summary>
    ///     Gets or sets whether the text output of R should be rendered instead of its graphics output.
    /// </summary>
    public bool RenderTextOutputInstead { get; set; }

    /// <summary>
    ///     Gets or sets the device's height in millimeters.
    /// </summary>
    public double DeviceHeightInMillimeters { get; set; }

    /// <summary>
    ///     Gets or sets the font in which text output should be rendered.
    /// </summary>
    public Font TextOutputFont { get; set; }

    /// <summary>
    ///     Gets or sets the brush in which text output should be rendered.
    /// </summary>
    public Brush TextOutputBrush { get; set; }

    /// <summary>
    ///     Gets or sets the device's resolution in dots-per-inch.
    /// </summary>
    public float DeviceResolutionInDpi { get; set; }

    /// <summary>
    ///     Runs the specified code and captures the device output in an <see cref="System.Drawing.Image" /> object.
    /// </summary>
    public Image GetDeviceOutputAsImage(string rCode, out string rScriptConsoleOutput, int timeoutMilliseconds = 1200000)
    {
        // notes: - the basic approach below is to have RScript.exe generate a PNG file which contains the graphics device output.
        //          the PNG file is then returned as Image object, and the remaining files are deleted (even if an exception is encountered).
        //        - although not impossible (via R.NET), we intentionally did not use an in-process approach here to generate the PNG files.
        //          using an out-of-process approach should make the code more stable.

        // return a blank image if rCode is empty
        if (rCode == null || rCode.Trim() == "")
        {
            rScriptConsoleOutput = "";
            var blankImage = new Bitmap(1, 1);
            blankImage.SetResolution(DeviceResolutionInDpi, DeviceResolutionInDpi);
            return blankImage;
        }

        // initialization
        string tempImageFileName = null;
        string tempScriptFileName = null;
        var rScriptConsoleOutputStringBuilder = new StringBuilder();

        try
        {
            // get a temp filename for the image file and the script file being run
            tempImageFileName = Path.GetTempFileName();
            tempScriptFileName = Path.GetTempFileName();

            // assemble the final R script and write it to a temp file (the code including all the wrapping code required to get the device output as PNG file)
            const string wrappingScriptTemplate = @"
                png(filename=""{0}"", width={1}, height={2}, units=""mm"", res={3});
                {4}
                dev.off();
            ";
            var finalRScript = string.Format(wrappingScriptTemplate, tempImageFileName.Replace(@"\", @"\\"),
                (int) DeviceWidthInMillimeters, (int) DeviceHeightInMillimeters, DeviceResolutionInDpi, rCode);
            File.WriteAllText(tempScriptFileName, finalRScript);

            // run RScript.exe and have it generate the PNG file
            // note: we use a timeout here to avoid that the process will run forever for whatever reasons.
            //       in case the timeout is reached, we kill the RScript.exe process and return with an exception
            var rScriptProcess = new Process
            {
                StartInfo =
                {
                    UseShellExecute = false,
                    ErrorDialog = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    FileName = Path.Combine(GetRBinPath(), "RScript.exe"),
                    Arguments = tempScriptFileName
                }
            };
            rScriptProcess.ErrorDataReceived += (sender, errorLine) =>
            {
                if (errorLine.Data != null)
                {
                    rScriptConsoleOutputStringBuilder.AppendLine(errorLine.Data);
                }
            };
            rScriptProcess.OutputDataReceived += (sender, outputLine) =>
            {
                if (outputLine.Data != null)
                {
                    rScriptConsoleOutputStringBuilder.AppendLine(outputLine.Data);
                }
            };
            rScriptProcess.Start();
            rScriptProcess.BeginErrorReadLine();
            rScriptProcess.BeginOutputReadLine();
            rScriptProcess.WaitForExit(timeoutMilliseconds);
            if (rScriptProcess.HasExited == false)
            {
                if (rScriptProcess.Responding)
                {
                    rScriptProcess.Close();
                    if (!rScriptProcess.HasExited)
                    {
                        rScriptProcess.Kill();
                    }
                }
                else
                {
                    rScriptProcess.Kill();
                }
            }

            // update rScriptConsoleOutput
            // note: this step is necessary because we cannot update out variables directly.
            //       also, it is better to use a StringBuilder for such type of strings.
            rScriptConsoleOutput = rScriptConsoleOutputStringBuilder.ToString();

            // check if RScript.exe returned successfully, ie. error code is 0
            if (rScriptProcess.ExitCode == 0)
            {
                // yes, successful return

                // grab the PNG file as Image object (incl. resolution information)
                Bitmap imageAsBitmap;
                using (var stream = new FileStream(tempImageFileName, FileMode.Open, FileAccess.Read))
                {
                    imageAsBitmap = new Bitmap(Image.FromStream(stream));
                    imageAsBitmap.SetResolution(DeviceResolutionInDpi, DeviceResolutionInDpi);
                }

                // return the image object if we are not to return the text output
                if (!RenderTextOutputInstead)
                {
                    return imageAsBitmap;
                }

                // if we reach here, we are to return the text output

                // create and return the text output
                var textOutputBitmap = imageAsBitmap;
                var textGraphics = Graphics.FromImage(textOutputBitmap);
                textGraphics.SmoothingMode = SmoothingMode.AntiAlias;
                textGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                textGraphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                textGraphics.DrawString(rScriptConsoleOutput, TextOutputFont, TextOutputBrush, new PointF(8f, 8f));
                textGraphics.Flush();
                return textOutputBitmap;
            }
            else
            {
                // no, something went wrong
                // throw an exception incl. the console output
                throw new Exception("RScript.exe returned with an unsuccessful return code. Console Output: " +
                                    rScriptConsoleOutput);
            }
        }
        finally
        {
            // if we reach here, we have either successfully run the code above or have encountered an exception.
            // in both cases we need to ensure that we leave none of the two files potentially created above.
            if (File.Exists(tempImageFileName) && tempImageFileName != null)
            {
                File.Delete(tempImageFileName);
            }
            if (File.Exists(tempScriptFileName) && tempScriptFileName != null)
            {
                File.Delete(tempScriptFileName);
            }
        }
    }

    /// <summary>
    ///     Returns the path of R's bin directory.
    /// </summary>
    /// <remarks>
    ///     Uses the Windows registry to find out about the version to be used. Ensure that the Windows registry has the
    ///     right version and install path set under Computer\HKEY_LOCAL_MACHINE\SOFTWARE\R-core.
    /// </remarks>
    private static string GetRBinPath()
    {
        // determine if we are 32 or 64 bit
        var is64Bit = IntPtr.Size == 8;

        // open the corresponding R-code key in the registry
        var rCoreKeyPath = @"SOFTWARE\R-core" + (is64Bit ? @"\R64" : @"\R");
        var rCoreKey = Registry.LocalMachine.OpenSubKey(rCoreKeyPath);
        if (rCoreKey == null)
        {
            throw new ApplicationException(
                string.Format(
                    "Registry key for R-core ({0}) not found. Expected key under '{1}'. Ensure that R is installed properly.",
                    is64Bit ? "64 bit" : "32 bit", rCoreKeyPath));
        }

        // get the Current Version and Install Path under the R-core\R / R-core\R64 key
        var currentVersion = new Version((string) rCoreKey.GetValue("Current Version"));
        var installPath = (string) rCoreKey.GetValue("InstallPath");

        // get the path of the bin subfolder in the Install Path
        var binPath = Path.Combine(installPath, "bin");

        // return the right bin path
        // note: up to 2.11.x, DLLs are installed in R_HOME\bin.
        //       from 2.12.0, DLLs are installed in the one level deeper directory.
        return currentVersion < new Version(2, 12) ? binPath : Path.Combine(binPath, is64Bit ? "x64" : "i386");
    }
}